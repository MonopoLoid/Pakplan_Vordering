Attribute VB_Name = "Module1"
Option Explicit
' =============================================================================
' MODULE:  Main Workbook Logic
' AUTHOR:  Tiaan (with help from Google, StackOverflow, and GitHub)
'
' Human-readable version shown in the ribbon (LocalConfig.GetVersionLabel).
' Bump this - and tag the matching commit, e.g. `git tag v1.9` - whenever
' you push an update that Git_Update should pull. Not used for the actual
' update-needed decision - that compares full file content instead.
Public Const APP_VERSION As String = "1.9"
'
' PURPOSE:
'   This module manages the "Pakplan Vordering" (Packing Plan Progress) system.
'   It pulls pallet intake and dispatch data from a SQL database (via Power Query),
'   builds a structured progress-tracking sheet ("Vordering"), calculates pallet
'   counts per variety/grade/pack/size, generates a summary sheet ("Opsomming"),
'   creates pie charts ("Grafieke"), and exports + emails a report workbook.
'
' SHEET OVERVIEW:
'   "Pakplan"    - Source packing plan (read from here, not modified much)
'   "Data"       - Power Query output: pallet intake from SQL DB
'   "Vordering"  - Main progress tracking sheet (built by Setup_Stuff)
'   "Opsomming"  - Summary sheet (built by Short_Stuff)
'   "Grafieke"   - Charts sheet (built by Chart_Stuff)
'
' MAIN ENTRY POINTS (also exposed to the Ribbon via _R wrapper subs):
'   Setup_Stuff  - Builds the Vordering sheet from Pakplan
'   Update_Stuff - Refreshes SQL data and recalculates all values
'   Input_Stuff  - Writes COUNTIFS formulas into Vordering cells
'   Short_Stuff  - Builds/updates the Opsomming summary sheet
'   Chart_Stuff  - Builds/updates the Grafieke charts sheet
'   Export_Stuff - Saves a clean copy of the workbook and emails it
'
' KEY CONSTANTS / GLOBALS:
'   startline    - The row on Vordering where data headers begin (row 3)
'   mServer/mDB  - SQL server connection strings (defined in a separate module)
'   mSave/mTo/mCC/mBCC - Export path and email addresses (separate module)
' =============================================================================

' --- Module-level variables ---
' These are shared across multiple subs in this module.

Dim curper As Integer       ' Tracks current progress bar % between sub calls
Dim totCol As Integer       ' Column number (as integer) of the TOTAL column in Vordering
Dim weeknumber As Integer   ' Week number extracted from the Pakplan title cell
Dim totPerc As Integer      ' Current overall progress percentage (0-100)
Dim vari As String          ' Variety name (read from Vordering cell D4)
Dim farmarray() As String   ' (Reserved) Array of farm codes - not fully implemented
Dim tName As String         ' Pack type name parsed from the Pakplan title
Dim pName As String         ' Sheet name for "Pakplan"
Dim oName As String         ' Sheet name for "Oorsig"
Dim shName As String        ' Sheet name for "Vordering"
Dim ansName As String       ' Sheet name for "Data" (the Power Query output sheet)
Dim WeekNumb As String      ' Week number string for file naming / email subject
Dim ribref As Boolean       ' True if Input_Stuff was called from the Ribbon directly
Dim pickedref1 As String    ' Starting Pick Reference for the data query range
Dim pickedref2 As String    ' Ending Pick Reference for the data query range
Dim dTable As ListObject    ' (Reserved) Reference to the DataQuery table object
Public gRibbon As IRibbonUI ' Ribbon UI reference for refreshing custom controls
Public gAbortPipeline As Boolean ' Set by HandleModuleError; checked by Update_Stuff
                                  ' between steps so one failure doesn't cascade into
                                  ' every later step also erroring and needing its own
                                  ' dialog acknowledged.

'Oorsig add-on
'Const OORSIG_NAME  As String = "Oorsig"
Const GAP_COLS     As Integer = 3
Const MAX_CMT_COLS As Integer = 10

' Row on Vordering/Pakplan where the actual data header row sits.
' Rows 1 and 2 are used for title/formatting; row 3 is the header.
Const startline As Integer = 3


' =============================================================================
' ENSURE DATA SHEET DEFAULTS
' Called for real from ThisWorkbook.Workbook_Open on every file open, and
' defensively re-checked here (via sheetExists) at the top of Update_Stuff
' in case the Data sheet was deleted, or macros were disabled at open and
' only enabled afterwards.
' Ensures the "Data" sheet exists and initialises key control cells used
' throughout the system (ribbon state, pick reference anchors, etc.).
' =============================================================================
Public Sub EnsureDataSheetDefaults()

    ' Create the Data sheet if it doesn't exist yet
    If Not sheetExists("Data") Then
        ThisWorkbook.Sheets.Add.name = "Data"
        ThisWorkbook.Worksheets("Data").Visible = xlSheetVisible
        ' NOTE: The table creation block below is commented out.
        ' If you need to re-enable it, uncomment and make sure
        ' "DataQuery" doesn't already exist before calling .Add.
        'If Not tableExists("Data", "DataQuery") Then
        '    Set dTable = ThisWorkbook.Worksheets("Data").ListObjects.Add(...)
        '    dTable.Name = "DataQuery"
        'End If
    End If

    ' Initialise ribbon/filter settings (if not yet set):
    '   FarmFilter = "All" / "Mahela" / "Other" - used in Power Query M code
    '   ValenciaGrouping = "ON" groups DEL/APV/MKN/GSV as "VAL"; "OFF" keeps separate
    '   WeekAutoMode = "ON" = use today's pick ref; "OFF" = user must enter
    If GetAppSetting("FarmFilter", "") = "" Then SetAppSetting "FarmFilter", "All"
    If GetAppSetting("ValenciaGrouping", "") = "" Then SetAppSetting "ValenciaGrouping", "OFF"
    If GetAppSetting("WeekAutoMode", "") = "" Then SetAppSetting "WeekAutoMode", "OFF"

    ' Calculate and store the current Pick Reference codes as PickRef1/PickRef2.
    ' Pick References are encoded as a 4-digit string combining week number and day.
    ' The encoding differs for week numbers < 10 vs >= 10 to keep them sortable.
    '
    ' PickRef1 = "Start of week" pick ref  (e.g. week 7 = "7100")
    ' PickRef2 = "Current day" pick ref    (e.g. week 7 Monday = "7100", Tuesday = "7200")
    '
    ' TODO: The magic number 35 in GetRainbowColor and the encoding logic here
    '       are interrelated. If the pick ref format changes, update both places.
    If WorksheetFunction.WeekNum(Date, vbSunday) < 10 Then
        ' Single-digit week: format is WeekNum & DayOfWeek & "00"
        SetAppSetting "PickRef2", WorksheetFunction.WeekNum(Date, vbMonday) & Weekday(Date, vbMonday) & "00"
        SetAppSetting "PickRef1", WorksheetFunction.WeekNum(Date, vbMonday) & "100"
    Else
        ' Double-digit week: split the digits so the string stays 4 chars and sortable
        ' e.g. week 12, day 3 => "2103" (last digit of week & "10" & first digit of week... CHECK THIS)
        ' TODO: This encoding is non-obvious. Consider a cleaner approach or add a unit test.
        SetAppSetting "PickRef2", Right(Str(WorksheetFunction.WeekNum(Date, vbMonday)), 1) & Weekday(Date, vbMonday) & "0" & Mid(Str(WorksheetFunction.WeekNum(Date, vbMonday)), 2, 1)
        SetAppSetting "PickRef1", Right(Str(WorksheetFunction.WeekNum(Date, vbMonday)), 1) & "10" & Mid(Str(WorksheetFunction.WeekNum(Date, vbMonday)), 2, 1)
    End If

End Sub


' =============================================================================
' HELPER: sheetExists
' Returns True if a sheet with the given name exists in the workbook.
' Used before adding new sheets or accessing sheets that may not exist yet.
'
' PARAMETERS:
'   sheetToFind  - Name of the sheet to look for
'   InWorkbook   - Optional: which workbook to check (defaults to ThisWorkbook)
' =============================================================================
Public Function sheetExists(sheetToFind As String, Optional InWorkbook As Workbook) As Boolean
    If InWorkbook Is Nothing Then Set InWorkbook = ThisWorkbook
    On Error Resume Next
    sheetExists = Not InWorkbook.Sheets(sheetToFind) Is Nothing
    On Error GoTo 0
End Function


' =============================================================================
' HELPER: tableExists
' Returns True if a ListObject (Excel Table) with the given name exists
' on the specified sheet. Used to avoid duplicate table creation.
'
' PARAMETERS:
'   sheetName  - Name of the worksheet to check
'   tableName  - Name of the ListObject/table to find
' =============================================================================
Function tableExists(sheetName As String, tableName As String) As Boolean
    Dim ws As Worksheet
    Dim tbl As ListObject
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    If ws Is Nothing Then Exit Function
    For Each tbl In ws.ListObjects
        If tbl.name = tableName Then
            tableExists = True
            Exit Function
        End If
    Next tbl
End Function


' =============================================================================
' HELPER: fnDateFromWeek
' Calculates a specific date from a year, ISO-ish week number, and weekday.
'
' NOT CURRENTLY CALLED FROM ANYWHERE in this module - kept for reference or
' possible future use. Flagging rather than removing since you're reading
' through to decide what's worth keeping.
'
' PARAMETERS:
'   iYear     - 4-digit year
'   iWeek     - Week number (1-53)
'   iWeekDday - Day within that week. The function calls Weekday() with no
'               override, so VBA's default applies: 1=Sunday, 2=Monday, ...
'               7=Saturday - verified below, not "1=Mon" as you might guess.
' =============================================================================
Function fnDateFromWeek(ByVal iYear As Integer, ByVal iWeek As Integer, ByVal iWeekDday As Integer)
    ' HOW: DateSerial(iYear, 1, N) means "day N counting from Jan 1 of iYear" -
    ' N doesn't have to be a valid day-of-January number, VBA just rolls the
    ' date forward that many days, which is what makes this one-liner work.
    '   (iWeek - 1) * 7          -> how many full weeks to skip before week iWeek
    '   + iWeekDday              -> then step to the target day within that week
    '   - Weekday(Jan1) + 1      -> correction so "day 1 of week 1" lines up with
    '                               whatever weekday Jan 1 actually falls on that
    '                               year (Jan 1 isn't always a Monday)
    ' Example: iYear=2026, iWeek=7, iWeekDday=1 (Sunday, per the default above).
    '   Jan 1, 2026 is a Thursday, so Weekday(Jan1)=5.
    '   (7-1)*7 + 1 - 5 + 1 = 42 + 1 - 5 + 1 = 39
    '   DateSerial(2026, 1, 39) rolls 39 days past Jan 1 -> 8 Feb 2026,
    '   which is indeed a Sunday, and the start of week 7.
    '   (Verified numerically, not just by hand - see if this matches your
    '   own understanding of which day the callers actually pass in.)
    fnDateFromWeek = DateSerial(iYear, 1, ((iWeek - 1) * 7) + iWeekDday - Weekday(DateSerial(iYear, 1, 1)) + 1)
End Function


' =============================================================================
' HELPER: charCheck
' Converts an ASCII character code to a column letter string.
' Handles columns beyond Z (e.g. AA, AB...) by prepending "A".
'
' PARAMETERS:
'   charVal - ASCII value of a column letter (65=A, 90=Z, 91+ = AA, AB...)
'
' NOTE: Currently only handles up to AZ (column 52). If you ever need
'       columns beyond AZ, this function needs to be expanded.
'       The current logic just prefixes "A" for anything > 90, which means
'       91="AA", 92="AB", ..., 116="AZ". Beyond that it would break.
' =============================================================================
Public Function charCheck(charVal) As String
    If charVal > 90 Then
        ' -26 wraps back into the A-Z range: 91 ("Z"+1) becomes 65 ("A") again,
        ' so charVal=91 -> Chr(65)="A" -> result "AA". charVal=92 -> Chr(66)="B"
        ' -> result "AB". It's re-using the same 65-90 range as a second "digit".
        charCheck = "A" & Chr(charVal - 26)
    Else
        charCheck = Chr(charVal)
    End If
End Function


' =============================================================================
' HELPER: CondFormAddRule
' Adds a conditional formatting rule to a range that colours the cell
' based on whether its value meets a numeric comparison against 0.
'
' PARAMETERS:
'   rng    - The Range to apply formatting to
'   sOp    - The comparison operator constant (e.g. 3=greater than, 5=less than, 6=equal)
'            XlFormatConditionOperator values:
'              xlBetween=1, xlNotBetween=2, xlEqual=3, xlNotEqual=4,
'              xlGreater=5, xlLess=6, xlGreaterEqual=7, xlLessEqual=8
'   lColor - The RGB colour to apply to the cell interior
' =============================================================================
Public Sub CondFormAddRule(rng, sOp, lColor)
    With rng.FormatConditions.Add(xlCellValue, sOp, "0")
        .Interior.Color = lColor
        .StopIfTrue = True  ' Stop evaluating further rules if this one matches
    End With
End Sub


' =============================================================================
' HELPER: ForceRibbonRefresh
' Forces the custom Ribbon labels (e.g. "Last Updated", "Last Sent")
' to re-read their values and redisplay.
' Called after data updates and after sending email.
' =============================================================================
Sub ForceRibbonRefresh()
    If Not gRibbon Is Nothing Then
        gRibbon.InvalidateControl "lblLastUpdated"
        gRibbon.InvalidateControl "lblLastSent"
    End If
End Sub


' =============================================================================
' HELPER: GetRainbowColor
' Returns an RGB Long colour that cycles through the rainbow spectrum.
' Used to animate the progress bar with a cycling colour effect.
'
' PARAMETERS:
'   progress - A number that increases as processing progresses.
'              The function wraps it within a 0-35 range (one full cycle = 35 units).
'
' NOTE: The value 35 is the cycle length. It was empirically chosen.
'       If you change the number of colour steps or the range, update this value.
'       The steps array defines: Red -> Orange -> Yellow -> Green -> Blue -> Violet -> Red
'
' WORKED EXAMPLE: progress=17 (out of the 0-35 cycle).
'   7 stops means UBound(steps)=6, so pos = 17/6 = 2.833.
'   i = Int(2.833) = 2  -> we're between steps(2)=Yellow and steps(3)=Green.
'   t = 2.833 - 2 = 0.833  -> 83.3% of the way from Yellow to Green.
'   r = 255 + 0.833*(0-255)   ~ 42
'   g = 255 + 0.833*(255-255)  = 255
'   b = 0   + 0.833*(0-0)      = 0
'   -> a yellow-green, which is exactly what you'd expect 83% of the way
'      from pure yellow toward pure green.
' =============================================================================
Function GetRainbowColor(progress As Double) As Long
    ' Define 7 colour stops for the full spectrum (the last repeats the first to close the loop)
    Dim steps As Variant
    steps = Array( _
        Array(255, 0, 0), _
        Array(255, 127, 0), _
        Array(255, 255, 0), _
        Array(0, 255, 0), _
        Array(0, 0, 255), _
        Array(198, 0, 198), _
        Array(255, 0, 0))
' Red
' Orange
' Yellow
' Green
' Blue
' Violet
' Red again (closes the loop)

    ' Wrap progress into the 0-35 range (one full rainbow cycle)
    Do
        If progress >= 35 Then progress = progress - 35
    Loop Until progress < 35

    ' Map progress (0-35) to a position along the colour stops (0-6)
    Dim pos As Double
    pos = progress / (UBound(steps))

    Dim i As Integer: i = Int(pos)          ' Which segment we're in
    Dim t As Double: t = pos - i            ' How far through that segment (0.0 to 1.0)

    ' Handle edge case: exactly at the last stop
    If i >= UBound(steps) Then
        GetRainbowColor = RGB(steps(i)(0), steps(i)(1), steps(i)(2))
        Exit Function
    End If

    ' Linearly interpolate between the two surrounding colour stops
    Dim r As Long, g As Long, b As Long
    r = steps(i)(0) + t * (steps(i + 1)(0) - steps(i)(0))
    g = steps(i)(1) + t * (steps(i + 1)(1) - steps(i)(1))
    b = steps(i)(2) + t * (steps(i + 1)(2) - steps(i)(2))

    GetRainbowColor = RGB(r, g, b)
End Function


' =============================================================================
' HELPER: GetProgressColor
' Returns a colour that transitions Red -> Yellow -> Green based on a 0.0-1.0
' progress fraction. Used to colour the progress bar during Input_Stuff and
' Short_Stuff (where rainbow cycling would be distracting).
'
' PARAMETERS:
'   progress - A value between 0.0 (start) and 1.0 (complete)
'              Values outside this range are clamped.
'
' MINOR QUIRK (found while documenting, not fixed - your call whether it's
' worth touching): the two phases aren't quite continuous at progress=0.5.
' Phase 1 approaches g=255*(0.5/0.5)=255 as progress nears 0.5 from below.
' Phase 2 at progress=0.5 exactly gives g=255-(0.5*255/8)=239.06. That's a
' visible ~16-unit dip in the green channel right at the halfway point -
' probably imperceptible during a fast-moving progress bar, but it's a real,
' verifiable discontinuity, not just a rounding artifact.
' =============================================================================
Function GetProgressColor(progress As Double) As Long
    If progress < 0 Then progress = 0
    If progress > 1 Then progress = 1

    Dim r As Long, g As Long, b As Long
    b = 0   ' No blue component in any phase

    If progress < 0.5 Then
        ' Phase 1 (0% to 50%): Red to Yellow
        ' Red stays at max; green ramps up from 0 to 255
        r = 255
        g = 255 * (progress / 0.5)
    Else
        ' Phase 2 (50% to 100%): Yellow to Green
        ' Red ramps down from 255 to 0.
        ' Green *should* stay at 255 for a pure Yellow->Green fade, but this
        ' formula ties it to `progress` itself (not the 0-1 phase-2 fraction),
        ' so it drifts down to 255-(1*255/8)=223 by the time progress=1.0 -
        ' a deliberate "slight fade to avoid overly bright green" per the
        ' original comment here, not a mistake, just worth knowing it's
        ' progress-linked rather than a fixed target colour.
        r = 255 * (1 - ((progress - 0.5) / 0.5))
        g = 255 - (progress * 255 / 8)
    End If

    GetProgressColor = RGB(r, g, b)
End Function


' =============================================================================
' HELPER: IsInSoftCitrusRange
' Checks whether a "soft citrus" size code (like "1X", "1XX") falls within
' a given low-to-high range using a custom ordering for these codes.
'
' Soft citrus sizes use a non-numeric ordering:
'   1XXXX > 1XXX > 1XX > 1X > 1 > 2 > 3 > 4 > 5 > 6  (largest to smallest)
' (i.e. 1XXXX is the largest fruit, 6 is the smallest)
'
' PARAMETERS:
'   val  - The size code to check (e.g. "1X", "2")
'   low  - The lower bound of the acceptable range (e.g. "1X")
'   high - The upper bound of the acceptable range (e.g. "3")
'
' RETURNS: True if val falls within [low, high] in the soft citrus ordering
'
' NOTE: 2-character codes that are NOT "1X", and 3-character codes that are
'       NOT "1XX" are immediately rejected (they don't belong to this system).
' =============================================================================
Function IsInSoftCitrusRange(val As String, low As String, high As String) As Boolean
    val = Replace(val, " ", "")  ' Strip spaces (Str() adds a leading space to numbers)

    ' Reject 2-char codes that aren't "1X"
    If Len(val) = 2 And val <> "1X" Then
        IsInSoftCitrusRange = False
        Exit Function
    End If
    ' Reject 3-char codes that aren't "1XX"
    If Len(val) = 3 And val <> "1XX" Then
        IsInSoftCitrusRange = False
        Exit Function
    End If

    ' Define the canonical ordering of soft citrus size codes
    Dim citrusOrder As Variant
    citrusOrder = Array("1XXXX", "1XXX", "1XX", "1X", "1", "2", "3", "4", "5", "6")

    Dim iVal As Long: iVal = -1
    Dim iLow As Long: iLow = -1
    Dim iHigh As Long: iHigh = -1
    Dim i As Long

    ' Find the array index for val, low, and high
    For i = LBound(citrusOrder) To UBound(citrusOrder)
        If citrusOrder(i) = val Then iVal = i
        If citrusOrder(i) = low Then iLow = i
        If citrusOrder(i) = high Then iHigh = i
    Next i

    ' Only return True if all three codes were found in the known order
    If iVal <> -1 And iLow <> -1 And iHigh <> -1 Then
        IsInSoftCitrusRange = (iVal >= iLow And iVal <= iHigh)
    Else
        IsInSoftCitrusRange = False  ' Unknown code = graceful fallback
    End If
End Function


' =============================================================================
' HELPER: IsCountAsOrdered
' Parses the comment text of a packing plan line to determine whether a
' given count/size should be treated as "packed as ordered" (i.e. count
' matches the order regardless of the standard size range).
'
' This is used in Setup_Stuff to choose the correct Overpack formula:
'   - "As Ordered" lines: overpack is measured against the actual ordered count
'   - Standard lines: overpack is measured against the total pallets needed
'
' PARAMETERS:
'   countToCheck - The size/count value from the Vordering header row (e.g. "64", "36(45)")
'   commentText  - The full comment string from the Pakplan COMMENTS column
'                  (e.g. "Pack 64 as ordered", "36-45 as ordered", "Pack as ordered")
'
' RETURNS: True if countToCheck falls within an "X as ordered" instruction
'          found in commentText.
'
' HOW IT WORKS:
'   1. If the comment says "PACK AS ORDERED" without a number prefix, all counts qualify.
'   2. Otherwise the regex finds patterns like "64ASORDERED", "36-45ASORDERED",
'      "56&64ASORDERED", etc. and checks if countToCheck falls in those ranges.
'   3. For compound counts like "36(45)", both parts are checked individually.
'   4. For soft citrus codes (1X, 1XX etc.), IsInSoftCitrusRange is used.
' =============================================================================
Function IsCountAsOrdered(countToCheck As String, commentText As String) As Boolean
    Dim pattern As String
    Dim matches As Object
    Dim regex As Object
    Dim part As Variant
    Dim brackCheck As Boolean, isNum As Boolean

    Set regex = CreateObject("VBScript.RegExp")
    commentText = Replace(UCase(commentText), " ", "")
    brackCheck = False
    isNum = False

    ' --- Special case: bare "PACK AS ORDERED" (no specific count prefix) ---
    If InStr(commentText, "PACKASORDERED") > 0 Then
        Dim packIndex As Long
        packIndex = InStr(commentText, "PACKASORDERED")
        ' Check whether there's a digit immediately before "PACKASORDERED"
        If packIndex > 1 Then isNum = IsNumeric(Mid(commentText, packIndex - 1, 1))
        If packIndex = 1 Or isNum = False Then
            ' No digit before it = applies to all counts
            IsCountAsOrdered = True
            Exit Function
        End If
    End If

    ' Strip "PACK" and "COUNT" keywords and commas before regex matching
    commentText = Replace(commentText, "PACK", "")
    commentText = Replace(commentText, "COUNT", "")
    commentText = Replace(commentText, ",", "")

    ' --- Handle compound counts: "36(45)" -> split into ["36", "45"] ---
    Dim countParts As Variant
    If InStr(countToCheck, "(") > 0 Then
        countParts = Split(Replace(Replace(countToCheck, "(", ","), ")", ""), ",")
        brackCheck = True
    Else
        ReDim countParts(0)
        countParts(0) = countToCheck
        brackCheck = False
    End If

    ' --- Regex: match patterns like "64ASORDERED", "36-45ASORDERED", "56&64ASORDERED" ---
    ' Pattern breakdown:
    '   ((\d+X*)(?:-(\d+X*))?)  = a count or range like "64", "36-45", "1X-3"
    '   (?:&(...))*              = optional additional counts joined by "&"
    '   ASORDERED                = literal end marker
    With regex
        .Global = True
        .IgnoreCase = True
        .pattern = "((\d+X*)(?:-(\d+X*))?)(?:&((\d+X*)(?:&(\d+X*))?))*ASORDERED"
    End With

    If regex.Test(commentText) Then
        Set matches = regex.Execute(commentText)
        Dim i As Integer
        For i = 0 To matches.Count - 1
            Dim m As Object: Set m = matches(i)

            ' Strip "ASORDERED" suffix, then split remaining by "&" to get each count/range
            Dim block As String: block = Replace(m, "ASORDERED", "")
            Dim segments As Variant: segments = Split(block, "&")
            Dim seg As Variant

            For Each seg In segments
                ' Each segment is either a single count ("64") or a range ("36-45")
                Dim lowerbound As String, upperbound As String
                If InStr(seg, "-") > 0 Then
                    lowerbound = Split(seg, "-")(0)
                    upperbound = Split(seg, "-")(1)
                Else
                    lowerbound = seg
                    upperbound = seg
                End If

                ' Check the count(s) against this range
                If brackCheck Then
                    ' Compound count: check each part separately
                    For Each part In countParts
                        If IsInSoftCitrusRange(Str(part), lowerbound, upperbound) Then
                            IsCountAsOrdered = True
                            Exit Function
                        ElseIf part >= lowerbound And part <= upperbound Then
                            IsCountAsOrdered = True
                            Exit Function
                        End If
                    Next part
                Else
                    ' Simple count: direct numeric or soft-citrus comparison
                    If IsInSoftCitrusRange(countToCheck, lowerbound, upperbound) Then
                        IsCountAsOrdered = True
                        Exit Function
                    ElseIf countToCheck >= lowerbound And countToCheck <= upperbound Then
                        IsCountAsOrdered = True
                        Exit Function
                    End If
                End If
            Next seg
        Next i
    End If
    ' If no match found, return False (implicit via unset boolean)
End Function

' =============================================================================
' RIBBON WRAPPER: Setup_Stuff_R
' Called by the Ribbon button. Delegates to Setup_Stuff.
' =============================================================================
Public Sub Setup_Stuff_R(control As IRibbonControl)
    Setup_Stuff
End Sub

' =============================================================================
' SETUP_STUFF
' Builds the "Vordering" (progress tracking) sheet from the "Pakplan" source.
'
' WHAT IT DOES:
'   1. Prompts the user if Vordering already exists (to confirm re-build).
'   2. Copies the header rows from Pakplan to Vordering.
'   3. For each packing line (rows where column M = "H" or "S"):
'      - Copies the line from Pakplan.
'      - Inserts 5 sub-rows: Pallets Needed/Outstanding/In Stock/Overpacked/Dispatched.
'      - Calculates cartons-per-pallet based on pack type and carton size.
'      - Writes formulas for each size column (SUM for totals, IFERROR for per-size).
'      - Writes OVERPACK formulas (with or without "As Ordered" logic).
'      - Adds conditional formatting (green/yellow/red) to key rows.
'   4. Adds a GRAND TOTAL row at the bottom.
'   5. Formats the sheet (borders, alignment, number format, row visibility).
'
' DEPENDENCIES:
'   - "Pakplan" sheet must exist and contain data starting with a "MAR*" header row.
'   - charCheck(), CondFormAddRule(), IsCountAsOrdered(), sheetExists()
'   - UserForm1 for progress display
'   - OptimizeVBA() to speed up processing
'
' KEY VARIABLES:
'   block       = 6: each Pakplan line expands to a 6-row block in Vordering
'   cartonCount = estimated cartons per pallet (varies by pack type and carton size)
'   totCol      = column index of the "TOTAL" column (module-level, set here)
'   stdcol      = column index of the "STD" column in Pakplan
' =============================================================================
Public Sub Setup_Stuff()
    On Error GoTo ErrHandler
    gAbortPipeline = False
    If Not EnsureEntitled() Then Exit Sub
    InitiateConstants

    ' --- Configure progress bar UserForm ---
    UserForm1.Width = 220
    UserForm1.Frame1.Width = 200
    UserForm1.Height = 98
    UserForm1.StartUpPosition = 2
    UserForm1.Caption = "Progress Bar"
    UserForm1.Label1.Caption = "0% Completed"
    UserForm1.Label2.Caption = ""
    UserForm1.Label3.Caption = "Setting things up..."
    UserForm1.Label2.Width = 0
    UserForm1.Label2.Height = UserForm1.Frame1.Height - 4
    UserForm1.Label2.BackColor = vbMagenta
    UserForm1.Frame1.Caption = ""

    Dim answer As String
    shName = "Vordering"
    pName = "Pakplan"
    answer = "6"   ' "6" = vbYes

    ' Ensure ribbon settings are initialised
    If GetAppSetting("FarmFilter", "") = "" Then SetAppSetting "FarmFilter", "All"
    If GetAppSetting("ValenciaGrouping", "") = "" Then SetAppSetting "ValenciaGrouping", "OFF"

    ' Create the Vordering sheet if it doesn't exist; otherwise ask the user
    If Not sheetExists(shName) Then
        ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).name = shName
        answer = "6"  ' Auto-proceed if sheet is new
    Else
        ThisWorkbook.Sheets(shName).Select
        answer = MsgBox("Do you wish to setup the " & shName & " sheet?", vbQuestion + vbYesNo, "User Response")
    End If

    If answer = "6" Then   ' User said Yes (or sheet was newly created)
        UserForm1.Show (False)
        OptimizeVBA (True)

        ' --- Clear existing data rows in Vordering (keep row 1 formatting) ---
        ThisWorkbook.Worksheets(shName).Rows(2 & ":" & ThisWorkbook.Worksheets(pName).Rows.Count).Delete

        ' --- Declare working variables ---
        Dim FirstCell As Range, LastCell As Range
        Dim rngTotalCol As Range, rngSTDCheck As Range
        Dim curRow As Integer, curCol As Integer, finalRow As Integer
        Dim h As Integer, i As Integer, j As Integer, k As Integer
        Dim l As Integer, m As Integer, n As Integer
        Dim block As Integer, cartonCount As Integer
        Dim totPerc As Integer, stdcol As Integer
        Dim colvar As String, commentCheck As String
        Dim sColstd As String, sColtot As String, sColterm As String
        Dim sColcomm As String, sColfinnum As String
        Dim sCol1 As String, sCol2 As String, sCol3 As String

        ' Find the extent of Pakplan data
        Set LastCell = ThisWorkbook.Worksheets(pName).Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        Set FirstCell = ThisWorkbook.Worksheets(pName).Cells.Find("MAR*", SearchOrder:=xlByRows, SearchDirection:=xlNext)
        finalRow = LastCell.Row



        ' Initialise counters
        commentCheck = ""
        colvar = ""
        curRow = 0: curCol = 0: block = 6: cartonCount = 0
        totCol = 0: totPerc = 0
        h = 0: i = 0: l = 0: k = 0: j = 0: m = 0

        ' --- Locate or insert the BATCH NR column in Pakplan ---
        ' The BATCH NR column should be immediately after the STD column.
        ' If it's missing, insert it (needed for Mahela vs non-Mahela logic).
        Set rngSTDCheck = ThisWorkbook.Sheets(pName).Range("I" & FirstCell.Row & ":Z" & FirstCell.Row).Find("STD", , xlValues, xlWhole)
        stdcol = rngSTDCheck.Column + 64
        sColstd = charCheck(stdcol + 1)
        If Not (UCase(ThisWorkbook.Sheets(pName).Range(sColstd & FirstCell.Row).Value) Like "BATC*") Then
            ThisWorkbook.Sheets(pName).Range(sColstd & ":" & sColstd).EntireColumn.Insert
            ThisWorkbook.Sheets(pName).Range(sColstd & FirstCell.Row).Value = "BATCH NR"
            ThisWorkbook.Sheets(pName).Columns(sColstd).ColumnWidth = 6
            ThisWorkbook.Sheets(pName).Range(sColstd & FirstCell.Row).WrapText = True
        End If

        ' --- Copy header rows from Pakplan to Vordering ---
        ' FINDING: "Range("A1").End(xlUp)" is a no-op here and everywhere else
        ' it appears in this loop (also lines below, and in the main loop's
        ' row-copy step). End(xlUp) searches UPWARD from the anchor cell for
        ' the last non-empty cell - starting from row 1, there's nowhere to
        ' go, so it always just returns A1 itself. This is the classic
        ' "find the last used row" idiom (normally written starting from the
        ' very BOTTOM of the sheet, e.g. Range("A1048576").End(xlUp)), just
        ' anchored at the top instead - so it never actually searches
        ' anything. Functionally identical to plain Range("A1").Offset(...),
        ' just harder to read. Not fixed here (this pass is about
        ' understanding, not changing behaviour) - flagging as a real
        ' candidate for a future cleanup pass.
        ThisWorkbook.Sheets(pName).Cells(1, "A").EntireRow.Copy Destination:=ThisWorkbook.Sheets(shName).Range("A" & 1).End(xlUp).Offset(startline - 3)
        ThisWorkbook.Sheets(pName).Range("A2:AZ2").Copy
        ThisWorkbook.Sheets(shName).Range("A2:AZ2").PasteSpecial xlPasteColumnWidths
        ThisWorkbook.Sheets(pName).Cells(FirstCell.Row, "A").EntireRow.Copy Destination:=ThisWorkbook.Sheets(shName).Range("A" & 1).End(xlUp).Offset(startline - 1)

        ' --- Find the TOTAL column in Vordering and derive related column references ---
        Set rngTotalCol = ThisWorkbook.Sheets(shName).Range("Q" & startline & ":AZ" & startline).Find("TOT*", , xlValues, xlWhole)
        totCol = rngTotalCol.Column + 64
        sColtot = charCheck(totCol)

        ' Hide the TOTAL column header text (small white font) to reduce clutter
        ThisWorkbook.Sheets(shName).Range(sColtot & (startline - 2)).Font.Color = vbWhite
        ThisWorkbook.Sheets(shName).Range(sColtot & (startline - 2)).Font.Size = 1

        ' Formatting for the data area
        ThisWorkbook.Sheets(shName).Range("C:" & sColtot).HorizontalAlignment = xlCenter
        ThisWorkbook.Sheets(shName).Range("Q:" & sColtot).NumberFormat = "0;-0;"  ' Suppress zeros

        ' COMMENTS column = 1 after TOTAL; second-to-last count column = 1 before TOTAL
        sColcomm = charCheck(totCol + 1)
        sColfinnum = charCheck(totCol - 1)

        ' =====================================================================
        ' MAIN LOOP: Process each row of Pakplan
        ' For each line where column M = "H" (High pallet) or "S" (Standard),
        ' copy it to Vordering and insert 5 sub-rows beneath it.
        ' =====================================================================
        For i = (startline) To finalRow

            ' j = calculated destination row offset in Vordering.
            ' Each matching line expands into `block` (6) rows, and k counts
            ' the non-matching (skipped) rows to subtract from the offset -
            ' this is what lets Vordering stay gapless even though Pakplan
            ' rows get skipped (comment rows, blank rows, etc. that aren't
            ' H or S). The actual destination row for the copied line itself
            ' is (1+j); each sub-row uses .Offset(j-i+N), which - because
            ' the "-i" and the later "+i" from Cells(i,...) cancel out -
            ' always lands on absolute row (j+N) regardless of what i is.
            '
            ' WORKED EXAMPLE (startline=3, block=6), verified numerically:
            '   i=3 (1st matching row), k=0: j=3  -> line at row 4,  sub-rows 5-9
            '   i=4 (2nd matching row), k=0: j=9  -> line at row 10, sub-rows 11-15
            '   i=5: skipped (not H/S)   -> k becomes 1
            '   i=6 (3rd matching row), k=1: j=15 -> line at row 16, sub-rows 17-21
            ' Notice rows 4-9, 10-15, 16-21 are perfectly back-to-back even
            ' though Pakplan row 5 was skipped in between - that's k doing its job.
            j = (i - (startline - 1)) * block - (k * block) - (block - startline)

            ' Only process rows flagged as H (Half pallet) or S (Standard)
            If (ThisWorkbook.Sheets(pName).Cells(i, "M").Value = "H") Or (ThisWorkbook.Sheets(pName).Cells(i, "M").Value = "S") Then

                ' Copy the Pakplan row to Vordering at the calculated offset
                ThisWorkbook.Sheets(pName).Cells(i, "A").EntireRow.Copy Destination:=ThisWorkbook.Sheets(shName).Range("A" & 1).End(xlUp).Offset(j)

                ' Draw a top border above this line's block
                sColterm = charCheck(totCol + 2)
                With ThisWorkbook.Sheets(shName).Range("A" & 1 & ":" & sColterm & 1).Offset(j).Borders(xlTop)
                    .LineStyle = xlContinuous
                    .Color = vbBlack
                    .Weight = xlMedium
                End With

                ' Update progress bar
                UserForm1.Label2.BackColor = GetRainbowColor(CDbl(i))
                totPerc = Round((i / (finalRow - 6)) * 100, 0)
                If totPerc > 100 Then totPerc = 100
                UserForm1.Label1.Caption = Str(totPerc) + "% Completed"
                UserForm1.Label2.Width = totPerc * 2
                UserForm1.Label3.Caption = "Setting things up..."
                DoEvents

                ' --- Write sub-row labels in column L ---
                With ThisWorkbook.Sheets(shName).Cells(i, "L")
                    .Offset(j - i + 2).Value = "Pallets Needed"
                    .Offset(j - i + 3).Value = "Pallets Outstanding"
                    .Offset(j - i + 4).Value = "Pallets in Stock"
                    .Offset(j - i + 5).Value = "Pallets Overpacked"
                    .Offset(j - i + 6).Value = "Pallets Dispatched"
                End With

                ' Record the actual Vordering row number for later formula building
                curRow = ThisWorkbook.Sheets(shName).Range("Q" & i).Offset(j - i + 1).Row

                ' --- Apply formatting to the 5 sub-rows (rows 2-6 of each block) ---
                For l = 2 To 6
                    With ThisWorkbook.Sheets(shName).Range("L" & i & ":O" & i).Offset(j - i + l)
                        .HorizontalAlignment = xlCenterAcrossSelection
                        .VerticalAlignment = xlCenter
                    End With
                    With ThisWorkbook.Sheets(shName).Range("A" & i & ":" & sColterm & i).Offset(j - i + l).Borders
                        .LineStyle = xlContinuous
                        .Color = vbBlack
                        .Weight = xlThin
                    End With
                    ' Grey background for all sub-rows
                    ThisWorkbook.Sheets(shName).Range("A" & i & ":" & sColterm & i).Offset(j - i + l).Interior.Color = RGB(210, 210, 210)
                Next l

                ' --- Copy TERM/COMMENTS text from Pakplan into Vordering sub-rows ---
                ' Searches backwards from the current row to find where the
                ' term text and comment text begin for this packing block.
                n = i + 1
                If Not ThisWorkbook.Sheets(pName).Cells(n - 1, 1).Value = "GR* TOT*" Then
                    Do
                        n = n - 1
                    Loop Until (Len(ThisWorkbook.Sheets(pName).Cells(n, (totCol - 64 + 2))) > 2) Or _
                               (ThisWorkbook.Sheets(pName).Cells(n, (totCol - 64 + 1)).Borders(xlEdgeTop).LineStyle <> xlNone)
                    ' Count how many term rows exist for this block
                    Dim counter As Integer
                    counter = 0
                    Do
                        counter = counter + 1
                    Loop Until (Len(ThisWorkbook.Sheets(pName).Cells(n + counter, (totCol - 64 + 2)).Value) > 2) Or _
                               (ThisWorkbook.Sheets(pName).Cells(n + counter, (totCol - 64 + 1)).Borders(xlEdgeTop).LineStyle <> xlNone) Or _
                               (ThisWorkbook.Sheets(pName).Cells(n + counter, 1).Value = "GRAND TOTAL") Or _
                               (counter > finalRow)
                    counter = counter - 1
                    ' Copy the term text rows
                    For l = 0 To counter
                        ThisWorkbook.Sheets(shName).Cells(i, (totCol - 64 + 1)).Offset(j - i + l + 1) = ThisWorkbook.Sheets(pName).Cells(n + l, (totCol - 64 + 1))
                        ThisWorkbook.Sheets(shName).Cells(i, (totCol - 64 + 1)).Offset(j - i + l + 1).Font.Underline = xlUnderlineStyleNone
                        ThisWorkbook.Sheets(shName).Cells(i, (totCol - 64 + 1)).Offset(j - i + l + 1).Font.Bold = False
                        ThisWorkbook.Sheets(shName).Cells(i, (totCol - 64 + 1)).Offset(j - i + l + 1).HorizontalAlignment = xlLeft
                    Next l
                End If

                ' --- Write SUM formulas into column P (total pallets per sub-row) ---
                ' Row 2 = Pallets Needed:    SUM of all size columns in that row
                ' Row 3 = Pallets Outstanding: same
                ' Row 4 = Pallets in Stock:   same
                ' Row 5 = Pallets Overpacked: IF "As Ordered" in comments -> special formula
                ' Row 6 = Pallets Dispatched: same as others
                For l = 2 To 6
                    curRow = ThisWorkbook.Sheets(shName).Range("P" & i).Offset(j - i + l).Row
                    ThisWorkbook.Sheets(shName).Range("P" & i).Offset(j - i + l).FormatConditions.Delete

                    If l <> 5 Then
                        ' Standard: sum all size columns for this row
                        ThisWorkbook.Sheets(shName).Range("P" & i).Offset(j - i + l).Formula = _
                            "=SUM(Q" & curRow & ":" & sColfinnum & curRow & ")"
                    Else
                        ' Pallets Overpacked: check if "AS ORDERED" appears in comments
                        ' If so, sum directly; otherwise derive from "Pallets Needed - Pallets Dispatched"
                        ThisWorkbook.Sheets(shName).Range("P" & i).Offset(j - i + l).Formula = _
                            "=IF(COUNTIF(" & sColcomm & (curRow - 4) & ":" & sColcomm & (curRow + 1) & "," & Chr(34) & "*AS ORDERED*" & Chr(34) & ")>0," & _
                            "SUM(Q" & curRow & ":" & sColfinnum & curRow & ")," & _
                            "IF(P" & (curRow - 2) & ">0,0,(-1)*P" & (curRow - 2) & "))"
                    End If

                    ' Conditional formatting on Outstanding (l=3) and Overpacked (l=5)
                    If l = 3 Or l = 5 Then
                        CondFormAddRule ThisWorkbook.Sheets(shName).Range("P" & i).Offset(j - i + l), 3, RGB(70, 170, 100)   ' = 0: green
                        CondFormAddRule ThisWorkbook.Sheets(shName).Range("P" & i).Offset(j - i + l), 6, RGB(230, 80, 80)    ' < 0: red
                        If l = 5 Then
                            CondFormAddRule ThisWorkbook.Sheets(shName).Range("P" & i).Offset(j - i + l), 5, RGB(230, 80, 80) ' > 0: red (overpack)
                        Else
                            CondFormAddRule ThisWorkbook.Sheets(shName).Range("P" & i).Offset(j - i + l), 5, vbYellow         ' > 0: yellow (still outstanding)
                        End If
                    End If
                Next l

                ' --- Calculate cartons per pallet for this line ---
                ' Preference order:
                '   1. Calculate from N+O columns (pallet dimensions)
                '   2. Hardcoded for specific pack types (Z10D, A06D)
                '   3. Lookup table by carton code, split by H (high) vs S (standard)
                curRow = ThisWorkbook.Sheets(shName).Range("Q" & i).Offset(j - i + 2).Row

                If ((ThisWorkbook.Sheets(shName).Cells(curRow - 1, "N").Value + ThisWorkbook.Sheets(shName).Cells(curRow - 1, "O").Value) > 0) And _
                   (InStr(1, ThisWorkbook.Sheets(shName).Cells(curRow - 1, "N").Value, "*") = 0) And _
                   (InStr(1, ThisWorkbook.Sheets(shName).Cells(curRow - 1, "O").Value, "*") = 0) Then
                    ' Use pallet dimension data if available and not asterisked
                    cartonCount = (ThisWorkbook.Sheets(shName).Cells(curRow - 1, sColtot).Value) / _
                                  (ThisWorkbook.Sheets(shName).Cells(curRow - 1, "N").Value + ThisWorkbook.Sheets(shName).Cells(curRow - 1, "O").Value)
                Else
                    If ThisWorkbook.Sheets(shName).Cells(curRow - 1, "H").Value = "Z10D" Or _
                       ThisWorkbook.Sheets(shName).Cells(curRow - 1, "H").Value = "A06D" Then
                        ' Special carton types always get 95
                        cartonCount = 95
                    Else
                        ' Look up by carton code. "H" = High pallet, other = standard pallet.
                        If ThisWorkbook.Sheets(shName).Cells(curRow - 1, "M").Value = "H" Then
                            ' High pallet carton counts
                            Select Case ThisWorkbook.Sheets(shName).Cells(curRow - 1, "H").Value
                                Case "A15C":         cartonCount = 80
                                Case "E10D", "E10D/D10D": cartonCount = 104
                                Case "D10D":         cartonCount = 112
                                Case "E15D", "E15C": cartonCount = 65
                                Case "D15D", "D15C": cartonCount = 70
                                Case "A07D":         cartonCount = 140
                                Case "G15C":         cartonCount = 50
                            End Select
                        Else
                            ' Standard pallet carton counts
                            Select Case ThisWorkbook.Sheets(shName).Cells(curRow - 1, "H").Value
                                Case "A15C":         cartonCount = 70
                                Case "E10D", "E10D/D10D": cartonCount = 88
                                Case "D10D":         cartonCount = 96
                                Case "E15D", "E15C": cartonCount = 55
                                Case "D15D", "D15C": cartonCount = 60
                                Case "A07D":         cartonCount = 120
                                Case "G15C":         cartonCount = 45
                            End Select
                        End If
                    End If
                End If

                ' --- Collect all comment text for this block (for IsCountAsOrdered check) ---
                commentCheck = ""
                For h = 1 To 6
                    commentCheck = commentCheck + ThisWorkbook.Sheets(shName).Range(sColcomm & i).Offset(j - i + h).Formula
                Next h

                ' --- Write per-size formulas for each size column ---
                ' Iterates from the first size column (curCol) to the last (totCol-1).
                curCol = ThisWorkbook.Sheets(shName).Range("Q" & i).Offset(j - i + 2).Column + 64

                For m = curCol To (totCol - 1)
                    colvar = charCheck(m)
                    curRow = ThisWorkbook.Sheets(shName).Range("Q" & i).Offset(j - i + 2).Row

                    ' Store cartons-per-pallet in the TERM column of the Dispatched row
                    ThisWorkbook.Sheets(shName).Range(sColterm & i).Offset(j - i + 6).Formula = cartonCount

                    ' Row 2 = Pallets Needed:
                    '   HOW THE CONDITION WORKS: InStr returns 0 when the search text
                    '   isn't found, or a positive position when it is. Adding four
                    '   InStr() calls together and checking "<> 0" is equivalent to
                    '   "is at least one of these found" (OR'd together) - it can only
                    '   sum to exactly 0 if ALL FOUR come back as 0.
                    '   WHAT IT'S ACTUALLY CHECKING: despite the comment below only
                    '   mentioning the asterisk, this also fires on "L", "M", or "S"
                    '   appearing anywhere in the cell - almost certainly the same
                    '   Large/Medium/Small size-override markers used elsewhere in the
                    '   codebase (see SizeMapping), meaning this cell holds a marker
                    '   or a compound value rather than a plain number either way.
                    '   If any of "*", "L", "M", "S" appear, this is a "split size"
                    '   line - add Stock + Dispatched from other rows instead.
                    If InStr(1, ThisWorkbook.Sheets(shName).Range(colvar & i).Offset(j - i + 1).Value, "*") + InStr(1, ThisWorkbook.Sheets(shName).Range(colvar & i).Offset(j - i + 1).Value, "L") + InStr(1, ThisWorkbook.Sheets(shName).Range(colvar & i).Offset(j - i + 1).Value, "M") + InStr(1, ThisWorkbook.Sheets(shName).Range(colvar & i).Offset(j - i + 1).Value, "S") <> 0 Then
                        ' Asterisk = use stock+dispatched sum instead (5.4.11.1 fix)
                        ThisWorkbook.Sheets(shName).Range(colvar & i).Offset(j - i + 2).Formula = _
                            "=IFERROR(" & colvar & (curRow + 2) & "+" & colvar & (curRow + 4) & ",0)"
                    Else
                        ThisWorkbook.Sheets(shName).Range(colvar & i).Offset(j - i + 2).Formula = _
                            "=IFERROR(" & colvar & (curRow - 1) & "/" & cartonCount & ",0)"
                    End If

                    ' Row 3 = Pallets Outstanding: Needed - (Stock + Dispatched)
                    curRow = curRow + 1
                    ThisWorkbook.Sheets(shName).Range(colvar & i).Offset(j - i + 3).Formula = _
                        "=IFERROR(" & colvar & (curRow - 1) & "-(" & colvar & (curRow + 1) & "+" & colvar & (curRow + 3) & "),0)"

                    ' Row 5 = Pallets Overpacked: formula differs by "As Ordered" logic
                    ' At this point curRow = the Overpacked row itself, so relative
                    ' to it: curRow-3=Needed, curRow-2=Outstanding, curRow-1=In Stock,
                    ' curRow+1=Dispatched (this is the same 6-row block, just addressed
                    ' from a different anchor point than earlier in this loop).
                    '
                    ' PLAIN-ENGLISH TRANSLATION of the "As Ordered" formula below:
                    '   IF Outstanding < 0 THEN -Outstanding
                    '   ELSE IF (InStock + Dispatched) - Needed > 0 THEN that difference
                    '   ELSE 0
                    ' Outstanding is itself "Needed - (InStock + Dispatched)" (see the
                    ' row-3 formula just above), so "Outstanding < 0" and
                    ' "(InStock+Dispatched) - Needed > 0" are algebraically the *same*
                    ' condition, just negated. That means once you reach the ELSE
                    ' branch (Outstanding >= 0), the inner ">0" check can never be
                    ' true - it always falls through to 0. Worth double-checking
                    ' against real data rather than taking my algebra on faith, but
                    ' if that holds, the inner IF is dead weight and this could
                    ' simplify to "IF(Outstanding<0, -Outstanding, 0)".
                    curRow = curRow + 2
                    If IsCountAsOrdered(ThisWorkbook.Sheets(shName).Range(colvar & startline).Formula, commentCheck) Then
                        ' "As Ordered": overpack = how much above the ordered count was packed/dispatched
                        ThisWorkbook.Sheets(shName).Range(colvar & i).Offset(j - i + 5).Formula = _
                            "=IF(" & colvar & (curRow - 2) & "<0,-" & colvar & (curRow - 2) & "," & _
                            "IF((" & colvar & (curRow - 1) & "+" & colvar & (curRow + 1) & ")-" & colvar & (curRow - 3) & ">0," & _
                            "(" & colvar & (curRow - 1) & "+" & colvar & (curRow + 1) & ")-" & colvar & (curRow - 3) & ",0))"
                    Else
                        ' Standard: overpack only if total pack+dispatch > total needed
                        ThisWorkbook.Sheets(shName).Range(colvar & i).Offset(j - i + 5).Formula = _
                            "=IF(($P" & (curRow - 1) & "+$P" & (curRow + 1) & ")>$P" & (curRow - 3) & "," & _
                            "IF((" & colvar & (curRow - 1) & "+" & colvar & (curRow + 1) & ")-" & colvar & (curRow - 3) & ">0," & _
                            "(" & colvar & (curRow - 1) & "+" & colvar & (curRow + 1) & ")-" & colvar & (curRow - 3) & ",0),0)"
                    End If
                Next m

                ' --- Add SUMIFS formulas for specific summary columns if this is a GRAND TOTAL block ---
                If ThisWorkbook.Sheets(shName).Range("A" & curRow - 1).Value = "GR* TOT*" Then
                    sCol2 = charCheck(curCol - 2)
                    sCol3 = charCheck(curCol - 3)
                    sCol1 = charCheck(curCol - 1)
                    ThisWorkbook.Sheets(shName).Range(sCol2 & i).Offset(j - i + 1).Formula = _
                        "=SUMIFS(" & sCol2 & "$1:" & sCol2 & "$" & (curRow - 2) & ",$" & sColtot & "$1:$" & sColtot & "$" & (curRow - 2) & "," & Chr(34) & ">0" & Chr(34) & ")"
                    ThisWorkbook.Sheets(shName).Range(sCol3 & i).Offset(j - i + 1).Formula = _
                        "=SUMIFS(" & sCol3 & "$1:" & sCol3 & "$" & (curRow - 2) & ",$" & sColtot & "$1:$" & sColtot & "$" & (curRow - 2) & "," & Chr(34) & ">0" & Chr(34) & ")"
                    ThisWorkbook.Sheets(shName).Range(sColtot & i).Offset(j - i + 1).Formula = _
                        "=SUMIFS(" & sColtot & "$1:" & sColtot & "$" & (curRow - 2) & ",$" & sColtot & "$1:$" & sColtot & "$" & (curRow - 2) & "," & Chr(34) & ">0" & Chr(34) & ")"
                    ThisWorkbook.Sheets(shName).Range(sCol1 & i).Offset(j - i + 5).Formula = _
                        "=SUMIFS(" & sCol1 & "$1:" & sCol1 & "$" & (curRow - 2) & ",$L$1:$L$" & (curRow - 2) & ",$L" & (curRow + 3) & ")"
                End If

            Else
                ' Row does not match H or S: skip it, increment the skip counter
                k = k + 1
            End If
        Next i

        ' =====================================================================
        ' GRAND TOTAL ROW
        ' Appended after all variety blocks. Sums each sub-row type across
        ' all varieties using SUMIFS filtered by the row label in column L.
        ' =====================================================================
        With ThisWorkbook.Worksheets(shName).Range("A1").CurrentRegion
            i = ThisWorkbook.Worksheets(shName).Cells(.Rows.Count, 1).Row + 1
        End With

        ' Format the 6 sub-rows of the Grand Total block
        For l = 1 To 6
            With ThisWorkbook.Sheets(shName).Range("L" & i & ":O" & i).Offset(j - i + l)
                .HorizontalAlignment = xlCenterAcrossSelection
                .VerticalAlignment = xlCenter
            End With
            With ThisWorkbook.Sheets(shName).Range("A" & i & ":" & sColterm & i).Offset(j - i + l).Borders
                .LineStyle = xlContinuous
                .Color = vbBlack
                .Weight = xlThin
            End With
            ThisWorkbook.Sheets(shName).Range("A" & i & ":" & sColterm & i).Offset(j - i + l).Font.Bold = True
            ThisWorkbook.Sheets(shName).Range("A" & i & ":" & sColterm & i).Offset(j - i + l).Interior.Color = RGB(165, 175, 200)
        Next l

        ' Write per-size SUMIFS formulas for the Grand Total sub-rows
        For m = curCol To (totCol - 1)
            colvar = charCheck(m)
            curRow = ThisWorkbook.Sheets(shName).Range("Q" & i).Offset(j - i + 2).Row
            For l = 2 To 6
                ThisWorkbook.Sheets(shName).Range(colvar & i).Offset(j - i + l).Formula = _
                    "=SUMIFS(" & colvar & "$1:" & colvar & "$" & (curRow - 2) & ",$L$1:$L$" & (curRow - 2) & ",$L" & (curRow - 1) & ")"
            Next l
            ThisWorkbook.Sheets(shName).Range(colvar & i).Offset(j - i + 1).Formula = _
                "=SUMIFS(" & colvar & "$1:" & colvar & "$" & (curRow - 2) & ",$" & sColtot & "$1:$" & sColtot & "$" & (curRow - 2) & "," & Chr(34) & ">0" & Chr(34) & ")"
        Next m

        ' Grand Total header label and top border
        ThisWorkbook.Sheets(shName).Range("A" & i).Offset(j - i + 1).Value = "GRAND TOTAL"
        With ThisWorkbook.Sheets(shName).Range("A" & i & ":" & sColterm & i).Offset(j - i + 1).Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Color = vbBlack
            .Weight = xlMedium
        End With

        ' Row labels for Grand Total sub-rows
        With ThisWorkbook.Sheets(shName).Cells(i, "L")
            .Offset(j - i + 2).Value = "Pallets Needed"
            .Offset(j - i + 3).Value = "Pallets Outstanding"
            .Offset(j - i + 4).Value = "Pallets in Stock"
            .Offset(j - i + 5).Value = "Pallets Overpacked"
            .Offset(j - i + 6).Value = "Pallets Dispatched"
        End With

        ' Grand Total per-row SUMIFS (summing by row label in column L)
        For m = curCol To (totCol - 1)
            colvar = charCheck(m)
            For l = 2 To 6
                ThisWorkbook.Sheets(shName).Range(colvar & i).Offset(j - i + l).Formula = _
                    "=SUMIFS(" & colvar & "$1:" & colvar & "$" & (curRow - 2) & ",$L$1:$L$" & (curRow - 2) & ",$L" & (curRow + l - 2) & ")"
            Next l
            ThisWorkbook.Sheets(shName).Range(colvar & i).Offset(j - i + 1).Formula = _
                "=SUMIFS(" & colvar & "$1:" & colvar & "$" & (curRow - 2) & ",$" & sColtot & "$1:$" & sColtot & "$" & (curRow - 2) & "," & Chr(34) & ">0" & Chr(34) & ")"
        Next m

        ' Summary column SUMIFS for Grand Total
        sCol2 = charCheck(curCol - 2)
        sCol3 = charCheck(curCol - 3)
        sCol1 = charCheck(curCol - 1)
        ThisWorkbook.Sheets(shName).Range(sCol2 & i).Offset(j - i + 1).Formula = _
            "=SUMIFS(" & sCol2 & "$1:" & sCol2 & "$" & (curRow - 2) & ",$" & sColtot & "$1:$" & sColtot & "$" & (curRow - 2) & "," & Chr(34) & ">0" & Chr(34) & ")"
        ThisWorkbook.Sheets(shName).Range(sCol3 & i).Offset(j - i + 1).Formula = _
            "=SUMIFS(" & sCol3 & "$1:" & sCol3 & "$" & (curRow - 2) & ",$" & sColtot & "$1:$" & sColtot & "$" & (curRow - 2) & "," & Chr(34) & ">0" & Chr(34) & ")"
        ThisWorkbook.Sheets(shName).Range(sColtot & i).Offset(j - i + 1).Formula = _
            "=SUMIFS(" & sColtot & "$1:" & sColtot & "$" & (curRow - 2) & ",$" & sColtot & "$1:$" & sColtot & "$" & (curRow - 2) & "," & Chr(34) & ">0" & Chr(34) & ")"
        ThisWorkbook.Sheets(shName).Range(sCol1 & i).Offset(j - i + 5).Formula = _
            "=SUMIFS(" & sCol1 & "$1:" & sCol1 & "$" & (curRow - 2) & ",$L$1:$L$" & (curRow - 2) & ",$L" & (curRow + 3) & ",$P$1:$P$" & (curRow - 2) & "," & Chr(34) & ">0" & Chr(34) & ")"

        ' Column P totals for Grand Total (sub-rows 2-6, skipping row 5 which is handled above)
        For l = 2 To 6
            If l <> 5 Then
                ThisWorkbook.Sheets(shName).Range(sCol1 & i).Offset(j - i + l).Formula = _
                    "=SUM(Q" & (curRow + l - 2) & ":" & sColfinnum & (curRow + l - 2) & ")"
            End If
        Next l

        ' --- Final formatting ---
        ThisWorkbook.Sheets(shName).Rows("1:" & j).EntireRow.Hidden = False
        ThisWorkbook.Sheets(shName).Range("P:P").NumberFormat = "General"
        ThisWorkbook.Sheets(shName).Range("A1").Select
        ThisWorkbook.Sheets(shName).Visible = xlSheetVisible
        Application.ScreenUpdating = True
        UserForm1.Hide

    End If  ' answer = "6"

    OptimizeVBA (False)
    Exit Sub
ErrHandler:
    HandleModuleError "Setup_Stuff"
End Sub


' =============================================================================
' RIBBON WRAPPER: Update_Stuff_R
' =============================================================================
Public Sub Update_Stuff_R(control As IRibbonControl)
    Update_Stuff
End Sub


' =============================================================================
' UPDATE_STUFF
' Master update routine. Refreshes SQL data via Power Query, then recalculates
' all pallet counts, summary, charts, and exports the file.
'
' FLOW:
'   1. Prompt user for Pick Reference range (start/end week-day codes).
'   2. Build Power Query M code for "Dispatches" query (out_det table).
'   3. Build Power Query M code for "DataQuery" (in_det table, filtered by
'      pick ref range, farm filter, and Valencia grouping setting).
'   4. Refresh both queries.
'   5. Call Input_Stuff  -> writes per-size COUNTIFS formulas into Vordering.
'   6. Call Short_Stuff  -> rebuilds the Opsomming summary sheet.
'   7. Call Chart_Stuff  -> rebuilds the Grafieke charts sheet.
'   8. Call Export_Stuff -> saves and emails the report workbook.
'
' PICK REFERENCE FORMAT:
'   Pick Refs are 4-digit codes used to filter the pallet intake data.
'   The encoding is: [last digit of week][day of week][0][first digit of week]
'   for double-digit weeks, or [week][day]["00"] for single-digit weeks.
'   These get rearranged into a sortable 4-char string: DDWW (day,day,week,week).
'   The Power Query uses a custom order list (CustomOrder) to filter by range.
'
' APP SETTINGS (CustomDocumentProperties - see LocalConfig.GetAppSetting):
'   FarmFilter       = "All" | "Mahela" | other (non-Mahela)
'   ValenciaGrouping = "ON" groups DEL/APV/MKN/GSV as "VAL"
'   WeekAutoMode     = "ON" auto-selects today's pick ref
'   PickRef1         = Last selected start pick ref
'   PickRef2         = Last selected end pick ref (current day)
' =============================================================================
Public Sub Update_Stuff()
    On Error GoTo ErrHandler
    gAbortPipeline = False
    If Not EnsureEntitled() Then Exit Sub
    InitiateConstants
    curper = 0

    Dim fullFile As String, currFile As String
    Dim wbName As String
    Dim wb As Workbook, wb1 As Workbook
    Dim i As Integer
    Dim farcurr() As String
    Dim farst() As Integer, faren() As Integer
    ReDim farcurr(18) As String
    ReDim farst(18) As Integer
    ReDim faren(18) As Integer
    totPerc = 0

    ' --- Configure progress bar ---
    UserForm1.Width = 220
    UserForm1.Frame1.Width = 200
    UserForm1.Height = 98
    UserForm1.StartUpPosition = 2
    UserForm1.Caption = "Progress Bar"
    UserForm1.Label1.Caption = "0% Completed"
    UserForm1.Label2.Caption = ""
    UserForm1.Label3.Caption = "Initiating update..."
    UserForm1.Label2.Width = 0
    UserForm1.Label2.Height = UserForm1.Frame1.Height - 4
    UserForm1.Label2.BackColor = vbRed
    UserForm1.Frame1.Caption = ""
    UserForm1.Show (False)

    OptimizeVBA (True)
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    shName = "Vordering"
    currFile = ThisWorkbook.FullName

    Dim disName As String
    ansName = "Data"
    disName = "V_Dispatches"

    ' Ensure the Data and Vordering sheets exist
    If Not sheetExists(ansName) Then
        EnsureDataSheetDefaults
    Else
        ThisWorkbook.Worksheets(ansName).Visible = xlSheetVisible
    End If
    If Not sheetExists(shName) Then
        Setup_Stuff
        If gAbortPipeline Then Exit Sub
    End If

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

    ' --- Declare variables for pick reference logic ---
    Dim dLastCell As Range, LastCell As Range
    Dim j As Integer, k As Integer, finalRow As Integer
    Dim weekval As Integer, weekval2 As Integer
    Dim checker As Integer, weeksnum As Integer, weekfnum As Integer
    Dim testvar As Integer, testvar2 As Integer, curPerc As Integer
    Dim pdchecker As String, pchecker As String
    Dim pcheck1 As String, pcheck2 As String, pcheck3 As String
    Dim curpickref As String, spickref As String
    Dim p1 As String, p2 As String, ans As String
    Dim found As Boolean, found2 As Boolean

    found = False: found2 = False
    weeksnum = 2: weekfnum = finalRow
    testvar = -1: testvar2 = -1
    p1 = GetAppSetting("PickRef1", "")   ' Previously saved start ref
    p2 = GetAppSetting("PickRef2", "")   ' Previously saved end ref
    ans = "7"   ' "7" = vbNo default

    ' --- Calculate today's pick reference strings ---
    ' curpickref = today's specific pick ref (week + day)
    ' spickref   = start-of-this-week pick ref (week + "100")
    If WorksheetFunction.WeekNum(Date, vbSunday) < 10 Then
        curpickref = WorksheetFunction.WeekNum(Date, vbMonday) & Weekday(Date, vbMonday) & "00"
        spickref = WorksheetFunction.WeekNum(Date, vbMonday) & "100"
    Else
        curpickref = Right(Str(WorksheetFunction.WeekNum(Date, vbMonday)), 1) & Weekday(Date, vbMonday) & "0" & Mid(Str(WorksheetFunction.WeekNum(Date, vbMonday)), 2, 1)
        spickref = Right(Str(WorksheetFunction.WeekNum(Date, vbMonday)), 1) & "10" & Mid(Str(WorksheetFunction.WeekNum(Date, vbMonday)), 2, 1)
    End If

    ' Block execution if Y1 is OFF and no saved pick ref exists
    If GetAppSetting("WeekAutoMode", "OFF") = "OFF" And p1 = "" Then
        MsgBox "Please at least select either 'Current Week' or provide a starting Pick Reference.", vbExclamation = vbOKOnly, "Select a Pick Ref."
        OptimizeVBA (False)
        UserForm1.Hide
        ThisWorkbook.Worksheets(shName).Activate
        Exit Sub
    End If

    pickedref1 = p1

    ' --- Determine pickedref1 (start of range) ---
    ' If Y1 = "OFF", use manual input or saved value; otherwise use today's week start
    ThisWorkbook.Activate
    If GetAppSetting("WeekAutoMode", "OFF") = "OFF" Then
        If p1 = "" Then
            ' Prompt user for start pick ref
            Do
                pickedref1 = InputBox("Enter starting Pick Reference needed: " & charCheck(13) & "(This week's Pick Reference is " & spickref & ")", "Pick Reference", spickref)
                If pickedref1 = "" Then pickedref1 = 0
                If Not IsNumeric(pickedref1) Then MsgBox "Please enter a valid number for the Pick Reference.", vbOKOnly
            Loop Until IsNumeric(pickedref1)
        End If
        ForceRibbonRefresh
    Else
        ' Auto mode: use this week's start pick ref
        pickedref1 = spickref
    End If

    ' --- Rearrange pick ref strings into sortable format ---
    ' HOW: "rotate so the 4th char goes first" means e.g. "2301" becomes
    ' "1230" (last char moves to the front, the rest shift right). Verified
    ' with real values: week=12, weekday=3 encodes as "2301" (see the
    ' week>=10 branch above: last-digit-of-week, weekday, "0",
    ' first-digit-of-week) - rotating gives "1230", putting the week's
    ' first digit at the front so string comparison sorts by week first.
    '
    ' FINDING (verified, not just suspected): single-digit weeks are NOT
    ' rotated the same way. week=9 encodes directly as "9300" (no rotation
    ' needed since it's already "week, day, 00"). Comparing week 9's "9300"
    ' against week 12's rotated "1230" as plain strings gives "9300" >
    ' "1230" - i.e. week 9 would sort AFTER week 12, backwards from actual
    ' chronological order. This only bites right at the week-9-to-10
    ' boundary of a season, so whether it's ever actually hit depends on
    ' whether your packing season spans that boundary while pick refs from
    ' both sides are being compared - worth checking against your own
    ' calendar rather than assuming either way. This is the same area the
    ' TODO further up already flagged as "non-obvious... consider a
    ' cleaner approach" - this finding is a concrete reason why.
    If Len(p1) < 4 Then p1 = "0" & p1
    If Len(p2) < 4 Then p2 = "0" & p2
    p1 = Right(p1, 1) & Left(p1, 3)
    p2 = Right(p2, 1) & Left(p2, 3)
    curpickref = Right(curpickref, 1) & Left(curpickref, 3)
    pickedref1 = Right(pickedref1, 1) & Left(pickedref1, 3)

    ' --- Determine pickedref2 (end of range) ---
    If GetAppSetting("WeekAutoMode", "OFF") = "OFF" Then
        ' Discard saved p2 if it's before p1
        If p2 < p1 Then p2 = ""
        If (p2 < curpickref) Then
            If p2 = "" Then
                p2 = curpickref
            Else
                ' Offer to use today's pick ref instead of the saved end ref
                p2 = Right(p2, 3) & Left(p2, 1)
                curpickref = Right(curpickref, 3) & Left(curpickref, 1)
                ans = MsgBox("Do you want to use the latest Pick Ref, " & curpickref & "?" & Chr(13) & "(Current selection is " & p2 & ".)", vbQuestion + vbYesNo, "User Response")
                p2 = Right(p2, 1) & Left(p2, 3)
                curpickref = Right(curpickref, 1) & Left(curpickref, 3)
            End If
        End If
        If ans = "6" Then p2 = curpickref  ' User said Yes

        ' If still no end ref, ask the user for one
        pickedref2 = Right(curpickref, 1) & Left(curpickref, 3)
        If p2 = "" Then
            Do
                If pickedref1 > curpickref Then
                    pickedref1 = Right(pickedref1, 3) & Left(pickedref1, 1)
                    pickedref2 = InputBox("Enter final Pick Reference needed: " & charCheck(13) & "(Selected start Pick Reference is " & pickedref1 & ")", "Week Number", pickedref1)
                    pickedref1 = Right(pickedref1, 1) & Left(pickedref1, 3)
                Else
                    curpickref = Right(curpickref, 3) & Left(curpickref, 1)
                    pickedref2 = InputBox("Enter final Pick Reference needed: " & charCheck(13) & "(Today's Pick Reference is " & curpickref & ")", "Week Number", curpickref)
                End If
                If pickedref2 = "" Then pickedref2 = 0
                pickedref2 = Right(pickedref2, 1) & Left(pickedref2, 3)
                If pickedref2 < pickedref1 Then pickedref2 = "Whoops"  ' Invalidate if end < start
                If Not IsNumeric(pickedref2) Then MsgBox "Please enter a valid number for the final Pick Reference number.", vbOKOnly
            Loop Until IsNumeric(pickedref2)
            pickedref2 = Right(pickedref2, 1) & Left(pickedref2, 3)
        Else
            pickedref2 = Right(p2, 3) & Left(p2, 1)
        End If
        pickedref2 = Right(pickedref2, 1) & Left(pickedref2, 3)
    Else
        ' Auto mode: use today as end ref, unless pickedref1 is in the future
        If pickedref1 > curpickref Then
            pickedref2 = pickedref1
        Else
            pickedref2 = curpickref
        End If
    End If

    ' --- Final format adjustment: ensure 4-digit pick refs with leading zeros ---
    ' Re-rotate from sortable back to original order, then zero-pad to 4 digits
    If (Mid(pickedref1, 2, 1) = "0") Then
        pickedref1 = Right(pickedref1, 2) & Left(pickedref1, 1)
    Else
        pickedref1 = Right(pickedref1, 3) & Left(pickedref1, 1)
    End If
    If (Mid(pickedref2, 2, 1) = "0") Then
        pickedref2 = Right(pickedref2, 2) & Left(pickedref2, 1)
    Else
        pickedref2 = Right(pickedref2, 3) & Left(pickedref2, 1)
    End If
    If Len(pickedref1) = 3 Then pickedref1 = "0" & pickedref1
    If Len(pickedref2) = 3 Then pickedref2 = "0" & pickedref2

    UserForm1.Show (False)

    ' --- Clear the Data sheet ready for fresh query results ---
    With ThisWorkbook.Worksheets(ansName)
        .Select
        On Error Resume Next
        .ShowAllData   ' Clear any active filters
        On Error GoTo ErrHandler
        If (.Rows.Count < 100000) Then
            .Rows(2 & ":" & .Rows.Count).Delete
        End If
        On Error Resume Next
        .ShowAllData
        On Error GoTo ErrHandler
    End With

    ' =========================================================================
    ' POWER QUERY: "Dispatches" (out_det table)
    ' Pulls PALLET_ID and SEQ_NO from the dispatch table.
    ' SEQ_NO > 0 means the pallet has been dispatched.
    ' =========================================================================
    Dim mCode As String
    Dim queryName As String

    queryName = "Dispatches"
    mCode = "let" & vbCrLf & _
        "    Source = Sql.Database(" & Chr(34) & mServer & Chr(34) & ", " & Chr(34) & mDB & Chr(34) & ")," & vbCrLf & _
        "    out_det_Sheet = Source{[Schema=" & Chr(34) & "dbo" & Chr(34) & ",Item=" & Chr(34) & "out_det" & Chr(34) & "]}[Data]," & vbCrLf & _
        "    #""Removed Other Columns"" = Table.SelectColumns(out_det_Sheet ,{""PALLET_ID"", ""SEQ_NO""})," & vbCrLf & _
        "    #""Changed Type"" = Table.TransformColumnTypes(#""Removed Other Columns"",{{""PALLET_ID"", type text}, {""SEQ_NO"", Int64.Type}})" & vbCrLf & _
        "in" & vbCrLf & _
        "#""Changed Type"""

    ' Update progress bar
    totPerc = 3: curper = totPerc
    UserForm1.Label1.Caption = Str(totPerc) + "% Completed"
    UserForm1.Label3.Caption = "Gathering Data..."
    UserForm1.Label2.Width = totPerc * 2
    DoEvents

    ThisWorkbook.Queries(queryName).Formula = mCode
    ThisWorkbook.Queries(queryName).Refresh

    ' Update progress bar
    totPerc = 4: curper = totPerc
    UserForm1.Label1.Caption = Str(totPerc) + "% Completed"
    UserForm1.Label3.Caption = "Transforming Data..."
    UserForm1.Label2.Width = totPerc * 2
    DoEvents

    ' =========================================================================
    ' POWER QUERY: "DataQuery" (in_det table)
    ' Pulls pallet intake records, trims whitespace, reorders columns,
    ' filters by the chosen pick reference range, optionally groups
    ' Valencia varieties, optionally filters to own-farm pallets,
    ' and left-joins with Dispatches to add a DISPATCH_MATCH column.
    '
    ' NOTE: The CustomOrder list defines the chronological order of all
    '       pick references for the season. If a pick ref is missing from
    '       this list, it will be excluded from results. Update the list
    '       at the start of each new season.
    '
    ' NOTE while documenting: OwnFarms below is a list of real farm
    ' identifier codes, and this file is pushed to the public repo. Lower
    ' sensitivity than the emails/server names already removed from
    ' LocalConfig (these codes don't obviously reveal who or where on
    ' their own), but flagging since it's the same category of "real
    ' operational data sitting in a public file" - your call whether it's
    ' worth moving to LocalConfig at some point (fetched into the M code
    ' as a variable instead of hardcoded here).
    ' =========================================================================
    queryName = "DataQuery"
    mCode = "let" & vbCrLf & _
        "    Source = Sql.Database(" & Chr(34) & mServer & Chr(34) & ", " & Chr(34) & mDB & Chr(34) & ")," & vbCrLf & _
        "    in_det_Sheet = Source{[Schema=" & Chr(34) & "dbo" & Chr(34) & ",Item=" & Chr(34) & "in_det" & Chr(34) & "]}[Data]," & vbCrLf & _
        "    #""Removed Other Columns"" = Table.SelectColumns(in_det_Sheet,{""PALLET_ID"", ""CONS_NO"", ""ORGZN"", ""VARIETY"", ""PACK"", ""GRADE"", ""MARK"", ""SIZE_COUNT"", ""INV_CODE"", ""PICK_REF"", ""FARM"", ""TARG_MKT"", ""REASON"", ""BATCH_NO"", ""TARGET_COUNTRY"", ""TARGET_REGION""})," & vbCrLf & _
        "    #""Changed Type"" = Table.TransformColumnTypes(#""Removed Other Columns"",{{""PALLET_ID"", type text}, {""CONS_NO"", type text}, {""ORGZN"", type text}, {""VARIETY"", type text}, {""PACK"", type text}, {""GRADE"", type text}, {""MARK"", type text}, {""SIZE_COUNT"", type text}, {""INV_CODE"", type text}, {""PICK_REF"", type text}, {""FARM"", type text}, {""TARG_MKT"", type text}, {""BATCH_NO"", type text}, {""TARGET_COUNTRY"", type text}, {""TARGET_REGION"", type text}})," & vbCrLf & _
        "    #""Trimmed Text"" = Table.TransformColumns(#""Changed Type"",{{""PALLET_ID"", Text.Trim, type text}, {""CONS_NO"", Text.Trim, type text}, {""ORGZN"", Text.Trim, type text}, {""VARIETY"", Text.Trim, type text}, {""PACK"", Text.Trim, type text}, {""GRADE"", Text.Trim, type text}, {""MARK"", Text.Trim, type text}, {""SIZE_COUNT"", Text.Trim, type text}, {""INV_CODE"", Text.Trim, type text}, {""PICK_REF"", Text.Trim, type text}, {""FARM"", Text.Trim, type text}, {""TARG_MKT"", Text.Trim, type text}, {""REASON"", Text.Trim, type text}, {""BATCH_NO"", Text.Trim, type text}, {""TARGET_COUNTRY"", Text.Trim, type text}, {""TARGET_REGION"", Text.Trim, type text}})," & vbCrLf & _
        "    #""Reordered Columns"" = Table.ReorderColumns(#""Trimmed Text"",{""PICK_REF"", ""CONS_NO"", ""ORGZN"", ""VARIETY"", ""TARG_MKT"", ""GRADE"", ""MARK"", ""PACK"", ""INV_CODE"", ""TARGET_REGION"", ""TARGET_COUNTRY"", ""BATCH_NO"", ""SIZE_COUNT"", ""FARM"", ""PALLET_ID"", ""REASON""})," & vbCrLf & _
        "    OwnFarms = {""D6613"", ""D6592"", ""D6612"", ""D9483"", ""D17440"", ""D6611"", ""D6814"", ""D6631"", ""D0534"", ""D0991"", ""D6595"", ""D0992"", ""D15976"", ""D17273"", ""D15563"", ""D15155"", ""D14034"", ""D13867""}," & vbCrLf & _
        "    CustomOrder = {""7100"", ""7200"", ""7300"", ""7400"", ""7500"", ""7600"", ""7700"", ""8100"", ""8200"", ""8300"", ""8400"", ""8500"", ""8600"", ""8700"", ""9100"", ""9200"", ""9300"", ""9400"", ""9500"", ""9600"", ""9700"", ""0101"", ""0201"", ""0301"", ""0401"", ""0501"", ""0601"", ""0701"", ""1101"", ""1201"", ""1301"", ""1401"", ""1501"", ""1601"", ""1701"", ""2101"", ""2201"", ""2301"", ""2401"", ""2501"", ""2601"", ""2701"", ""3101"", ""3201"", ""3301"", ""3401"", ""3501"", ""3601"", ""3701"", ""4101"", ""4201"", ""4301"", ""4401"", ""4501"", ""4601"", ""4701"", ""5101"", ""5201"", ""5301"", ""5401"", ""5501"", ""5601"", ""5701"", ""6101"", ""6201"", ""6301"", ""6401"", ""6501"", ""6601"", ""6701"", ""7101"", ""7201"", ""7301"", ""7401"", ""7501"", ""7601"", ""7701"", ""8101"", ""8201"", ""8301"", ""8401"", ""8501"", ""8601"", ""8701"", ""9101"", ""9201"", ""9301"", ""9401"", ""9501"", ""9601"", ""9701"", ""0102"", ""0202"", ""0302"", ""0402"", ""0502"", ""0602"", ""0702"", " & Chr(34) & _
        "1102"", ""1202"",""1302"", ""1402"", ""1502"", ""1602"", ""1702"", ""2102"", ""2202"", ""2302"", ""2402"", ""2502"", ""2602"", ""2702"", ""3102"", ""3202"", ""3302"", ""3402"", ""3502"", ""3602"", ""3702"", ""4102"", ""4202"", ""4302"", ""4402"", ""4502"", ""4602"", ""4702"", ""5102"", ""5202"", ""5302"", ""5402"", ""5502"", ""5602"", ""5702"", ""6102"", ""6202"", ""6302"", ""6402"", ""6502"", ""6602"", ""6702"", ""7102"", ""7202"", ""7302"", ""7402"", ""7502"", ""7602"", ""7702"", ""8102"", ""8202"", ""8302"", ""8402"", ""8502"", ""8602"", ""8702"", ""9102"", ""9202"", ""9302"", ""9402"", ""9502"", ""9602"", ""9702"", ""0103"", ""0203"", ""0303"", ""0403"", ""0503"", ""0603"", ""0703"", ""1103"", ""1203"", ""1303"", ""1403"", ""1503"", ""1603"", ""1703"", ""2103"", ""2203"", ""2303"", ""2403"", ""2503"", ""2603"", ""2703"", " & Chr(34) & _
        "3103"", ""3203"", ""3303"", ""3403"", ""3503"", ""3603"", ""3703"", ""4103"", ""4203"", ""4303"", ""4403"", ""4503"", ""4603"", ""4703"", ""5103"", ""5203"", ""5303"", ""5403"", ""5503"", ""5603"", ""5703"", ""6103"", ""6203"", ""6303"", ""6403"", ""6503"", ""6603"", ""6703"", ""7103"", ""7203"", ""7303"", ""7403"", ""7503"", ""7603"", ""7703"", ""8103"", ""8203"", ""8303"", ""8403"", ""8503"", ""8603"", ""8703"", ""9103"", ""9203"", ""9303"", ""9403"", ""9503"", ""9603"", ""9703""}," & vbCrLf & _
        "    StartValue = " & Chr(34) & pickedref1 & Chr(34) & "," & vbCrLf & _
        "    EndValue = " & Chr(34) & pickedref2 & Chr(34) & "," & vbCrLf & _
        "    StartIndex = List.PositionOf(CustomOrder, StartValue)," & vbCrLf & _
        "    EndIndex = List.PositionOf(CustomOrder, EndValue)," & vbCrLf & _
        "    #""Added Sort Index"" = Table.AddColumn(#""Reordered Columns"", ""SortIndex"", each List.PositionOf(CustomOrder, [PICK_REF]), Int64.Type)," & vbCrLf & _
        "    #""Filtered Range"" = Table.SelectRows(#""Added Sort Index"", each [SortIndex] >= StartIndex and [SortIndex] <= EndIndex)," & vbCrLf & _
        "    #""Removed SortIndex"" = Table.RemoveColumns(#""Filtered Range"", {""SortIndex""})," & vbCrLf

    ' --- Conditional M code: Valencia variety grouping ---
    ' ValenciaGrouping = "ON": group DEL/APV/MKN/GSV all under "VAL"
    ' ValenciaGrouping = "OFF": keep each variety separate
    If GetAppSetting("ValenciaGrouping", "OFF") = "ON" Then
        mCode = mCode & "    #""Grouped Valencia Varieties"" = Table.AddColumn(#""Removed SortIndex"", ""VARIETY_GROUP"", each if List.Contains({""DEL"", ""APV"", ""MKN"", ""GSV""}, [VARIETY]) then ""VAL"" else [VARIETY])," & vbCrLf
    Else
        mCode = mCode & "    #""Grouped Valencia Varieties"" = Table.AddColumn(#""Removed SortIndex"", ""VARIETY_GROUP"", each [VARIETY])," & vbCrLf
    End If

    ' --- Conditional M code: Farm filter ---
    ' FarmFilter = "Mahela":    only own-farm pallets (FARM in OwnFarms list)
    ' FarmFilter = "All":       all pallets regardless of farm
    ' FarmFilter = "NonMahela": non-own-farm pallets only
    If GetAppSetting("FarmFilter", "All") = "Mahela" Then
        mCode = mCode & "    #""Filtered Own Farms"" = Table.SelectRows(#""Grouped Valencia Varieties"", each List.Contains(OwnFarms, [FARM]))," & vbCrLf & _
            "    #""Merged Dispatch"" = Table.NestedJoin(#""Filtered Own Farms"", {""PALLET_ID""}, Dispatches, {""PALLET_ID""}, ""DispatchLookup"", JoinKind.LeftOuter)," & vbCrLf
    ElseIf GetAppSetting("FarmFilter", "All") = "All" Then
        mCode = mCode & "    #""Merged Dispatch"" = Table.NestedJoin(#""Grouped Valencia Varieties"", {""PALLET_ID""}, Dispatches, {""PALLET_ID""}, ""DispatchLookup"", JoinKind.LeftOuter)," & vbCrLf
    Else
        ' Covers "NonMahela" plus any unrecognized value as a safe default,
        ' so an unexpected FarmFilter can't silently produce incomplete M code.
        mCode = mCode & "    #""Filtered Own Farms"" = Table.SelectRows(#""Grouped Valencia Varieties"", each not List.Contains(OwnFarms, [FARM]))," & vbCrLf & _
            "    #""Merged Dispatch"" = Table.NestedJoin(#""Filtered Own Farms"", {""PALLET_ID""}, Dispatches, {""PALLET_ID""}, ""DispatchLookup"", JoinKind.LeftOuter)," & vbCrLf
    End If

    ' Final M code: expand the dispatch join, replace nulls, reorder columns
    mCode = mCode & _
        "    #""Expand Dispatch"" = Table.ExpandTableColumn(#""Merged Dispatch"", ""DispatchLookup"", {""SEQ_NO""}, {""DISPATCH_MATCH""})," & vbCrLf & _
        "    #""Dispatch Matched"" = Table.ReplaceValue(#""Expand Dispatch"", null, 0, Replacer.ReplaceValue, {""DISPATCH_MATCH""})," & vbCrLf & _
        "    #""Reorder Columns"" = Table.ReorderColumns(#""Dispatch Matched"",{""PICK_REF"", ""CONS_NO"", ""ORGZN"", ""VARIETY"", ""TARG_MKT"", ""GRADE"", ""MARK"", ""PACK"", ""INV_CODE"", ""TARGET_REGION"", ""TARGET_COUNTRY"", ""BATCH_NO"", ""SIZE_COUNT"", ""FARM"", ""PALLET_ID"", ""VARIETY_GROUP"", ""DISPATCH_MATCH"", ""REASON""})" & vbCrLf & _
        "in" & vbCrLf & _
        "    #""Reorder Columns"""

    ' Update progress bar
    totPerc = 5: curper = totPerc
    UserForm1.Label1.Caption = Str(totPerc) + "% Completed"
    UserForm1.Label3.Caption = "Gathering Data..."
    UserForm1.Label2.Width = totPerc * 2
    DoEvents

    ' Refresh the DataQuery (this loads data from SQL into the Data sheet)
    OptimizeVBA (False)
    ThisWorkbook.Queries(queryName).Formula = mCode
    ThisWorkbook.Queries(queryName).Refresh
    DoEvents
    OptimizeVBA (True)

    ' Update progress bar
    totPerc = 6: curper = totPerc
    UserForm1.Label1.Caption = Str(totPerc) + "% Completed"
    UserForm1.Label3.Caption = "Transforming Data..."
    UserForm1.Label2.Width = totPerc * 2
    DoEvents

    OptimizeVBA (False)
    ThisWorkbook.Activate
    ThisWorkbook.Worksheets(shName).Select
    OptimizeVBA (True)

    curper = totPerc
    ThisWorkbook.Worksheets(ansName).Select

    ' --- Step 5: Write COUNTIFS formulas into Vordering ---
    ribref = False   ' Called from Update, not directly from Ribbon
    Input_Stuff
    If gAbortPipeline Then Exit Sub

    ' --- Finish progress bar ---
    totPerc = 0
    UserForm1.Label1.Caption = "100% Completed"
    UserForm1.Label2.Width = 200
    UserForm1.Label3.Caption = "Finishing up"
    DoEvents

    OptimizeVBA (True)
    DoEvents

    ' --- Step 6: Rebuild Opsomming summary sheet ---
    Short_Stuff
    If gAbortPipeline Then Exit Sub

    OptimizeVBA (False)
    SetAppSetting "LastUpdated", "Last Updated: " & Now()
    DoEvents
    ForceRibbonRefresh

    ' --- Step 7: Rebuild Grafieke charts sheet ---
    Chart_Stuff
    If gAbortPipeline Then Exit Sub

    ThisWorkbook.Worksheets(shName).Visible = xlSheetVisible
    ThisWorkbook.Worksheets(shName).Select

    ' --- Step 8: Export and email the report ---
    Export_Stuff
    If gAbortPipeline Then Exit Sub

    OptimizeVBA (False)
    ThisWorkbook.Worksheets(shName).Visible = xlSheetVisible
    ThisWorkbook.Worksheets(shName).Select
    UserForm1.Hide

    Exit Sub
ErrHandler:
    HandleModuleError "Update_Stuff"
End Sub


' =============================================================================
' RIBBON WRAPPER: Input_Stuff_R
' Called directly from Ribbon. Sets ribref=True so Input_Stuff shows its own
' progress bar and hides it when done.
' =============================================================================
Sub Input_Stuff_R(control As IRibbonControl)
    ribref = True
    totPerc = 0
    Input_Stuff
End Sub


' =============================================================================
' INPUT_STUFF
' Writes COUNTIFS formulas into the "Pallets in Stock" and "Pallets Dispatched"
' rows of Vordering, looking up values from the Data sheet.
'
' WHAT IT DOES:
'   For each row in Vordering that is labelled "Pallets Dispatched" or
'   "Pallets in Stock", writes a COUNTIFS formula that counts matching pallets
'   in the Data sheet based on:
'     - Variety (col C), Variety Group (col P/D), Target Market (E), Grade (F),
'       Mark (G), Pack (H), Inv Code (I), Targ Mkt (J), Batch (K),
'       Dispatch Match (R), Farm (L), Size/Count (M)
'
'   The Size/Count lookup is the most complex part: some column headers in
'   Vordering use combined sizes like "36(45)" or "20,22" that need to be
'   mapped to the actual SIZE_COUNT values in the database.
'   The mapping differs by carton type (A15C vs others) and fruit type
'   (soft citrus uses X-codes like "1X", "1XX").
'
'   After writing formulas, a second pass checks for "doubled" lines
'   (same variety/grade/pack appearing twice in the Pakplan due to split
'   batches) and subtracts their counts to avoid double-counting.
'
' NOTE:
'   - `ribref = True`  means this sub was called standalone (shows its own
'                      progress bar and hides it when done)
'   - `ribref = False` means it was called from Update_Stuff (shares the bar)
' =============================================================================
Sub Input_Stuff()
    On Error GoTo ErrHandler
    gAbortPipeline = False
    If Not EnsureEntitled() Then Exit Sub

    Dim prog As Double
    Dim ansName As String
    Dim shName As String
    Dim totPerc As Integer
    Dim dval As Integer, ivalue As Integer
    Dim sForm As String, sPack As String
    ansName = "Data"
    shName = "Vordering"

    Dim i As Integer, j As Integer, k As Integer
    Dim vals As Integer, vorRows As Integer, totCol As Integer
    vals = 0
    totCol = ThisWorkbook.Sheets(shName).Range("Q" & startline & ":AZ" & startline).Find("TOTAL", , xlValues, xlWhole).Column
    vorRows = ThisWorkbook.Worksheets(shName).Columns(totCol).Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row + 5

    ' --- Pre-style: colour the Dispatched and In Stock row backgrounds ---
    For i = (vorRows - 4) To vorRows
        If (ThisWorkbook.Worksheets(shName).Cells(i, 12).Value = "Pallets Dispatched") Or _
           (ThisWorkbook.Worksheets(shName).Cells(i, 12).Value = "Pallets in Stock") Then
            ThisWorkbook.Sheets(shName).Cells(i, 16).Interior.Color = RGB(214, 220, 228)
            For j = 17 To totCol - 1
                ThisWorkbook.Sheets(shName).Cells(i, j).Interior.Color = RGB(214, 220, 228)
            Next j
        End If
    Next i

    vari = ThisWorkbook.Sheets(shName).Cells(4, 4).Value

    Dim doubledCheck As Integer
    doubledCheck = 0

    ' Show progress bar if called standalone from Ribbon
    If ribref Then
        UserForm1.Width = 220
        UserForm1.Frame1.Width = 200
        UserForm1.Height = 98
        UserForm1.StartUpPosition = 2
        UserForm1.Caption = "Progress Bar"
        UserForm1.Label1.Caption = "0% Completed"
        UserForm1.Label2.Caption = ""
        UserForm1.Label3.Caption = "Updating values..."
        UserForm1.Label2.Width = 0
        UserForm1.Label2.Height = UserForm1.Frame1.Height - 4
        UserForm1.Label2.BackColor = vbRed
        UserForm1.Frame1.Caption = ""
        UserForm1.Show (False)
    End If

    ' =========================================================================
    ' MAIN LOOP: For each row in Vordering, check if it's a count row.
    ' - "Pallets Dispatched" rows: count pallets that went through the packhouse
    '   but are NOT at packhouse (i.e. DISPATCH_MATCH = non-null -> dispatched)
    ' - "Pallets in Stock" rows: count pallets still AT packhouse
    ' =========================================================================
    For i = 1 To (vorRows - 5)

        ' Update progress bar colour and percentage
        prog = (i - 1) / (vorRows - 5 - 1)
        UserForm1.Label2.BackColor = GetProgressColor(CDbl(prog))
        If ribref Then
            totPerc = Round((i / (vorRows - 5)) * 100, 0)
        Else
            ' When called from Update_Stuff, progress starts from curper (already ~6%)
            totPerc = curper + Round((i / (vorRows - 5)) * 88, 0)
        End If
        If totPerc > 100 Then totPerc = 100
        UserForm1.Label1.Caption = Str(totPerc) + "% Completed"
        UserForm1.Label2.Width = totPerc * 2
        UserForm1.Label3.Caption = "Updating Values..."
        DoEvents

        ' Determine if this row is a counting row and set sPack filter accordingly
        ' vals = the row number of the corresponding Pakplan data row
        If ThisWorkbook.Worksheets(shName).Cells(i, 12).Value = "Pallets Dispatched" Then
            vals = i - 5        ' Data row is 5 rows above the Dispatched label row
            sPack = "<>PACKHOUSE"   ' Dispatched pallets are NOT at packhouse
        ElseIf ThisWorkbook.Worksheets(shName).Cells(i, 12).Value = "Pallets in Stock" Then
            vals = i - 3        ' Data row is 3 rows above the In Stock label row
            sPack = "PACKHOUSE"     ' In-stock pallets ARE at packhouse
        Else
            vals = 0            ' Not a counting row; skip
        End If

        If vals <> 0 Then
            ThisWorkbook.Sheets(shName).Cells(i, 16).Interior.Color = RGB(230, 230, 230)

            ' --- Build COUNTIFS formula for each size column ---
            For j = 17 To totCol - 1

                sForm = BuildFullFormula(shName, ansName, startline, j, vals, sPack, , totCol)

                ' Colour the cell light grey (will be overwritten if formula fills in)
                ThisWorkbook.Sheets(shName).Cells(i, j).Interior.Color = RGB(230, 230, 230)


                ' ---------------------------------------------------------------
                ' DOUBLED LINE CHECK (immediate neighbour)
                ' If the next block (6 rows down) has identical attributes but
                ' a different mark code ("D" prefix), subtract its dispatched
                ' count to avoid double-counting pallets that span two pick runs.
                ' ---------------------------------------------------------------
                With ThisWorkbook.Worksheets(shName)
                    ivalue = i - 3
                    dval = i + 3
                    If (dval < (vorRows - 5)) And _
                       (.Cells(dval, 3).Value = .Cells(ivalue, 3).Value) And _
                       (.Cells(dval, 5).Value = .Cells(ivalue, 5).Value) And _
                       (.Cells(dval, 6).Value = .Cells(ivalue, 6).Value) And _
                       (.Cells(dval, 7).Value = .Cells(ivalue, 7).Value) And _
                       (.Cells(dval, 8).Value = .Cells(ivalue, 8).Value) And _
                       (.Cells(dval, 9).Value = .Cells(ivalue, 9).Value) And _
                       (.Cells(dval, 10).Value = .Cells(ivalue, 10).Value) And _
                       (.Cells(dval, 11).Value = .Cells(ivalue, 11).Value) And _
                       (.Cells(dval, 16).Value = "D" & .Cells(ivalue, 16).Value) And _
                       (Not IsEmpty(.Cells(ivalue, 3))) Then
                        ' charCheck(j+64) turns the column NUMBER back into its own
                        ' letter (j is always 17-26 here, i.e. Q-Z, well inside the
                        ' plain Chr(64+n) trick charCheck falls back to for single
                        ' letters - see charCheck's own comment). So this appends
                        ' "-<sameCol><dval+3>-<sameCol><dval+5>-0" to the formula
                        ' already sitting in that cell: subtract the neighbour
                        ' block's Stock and Dispatched cells, then a literal "-0"
                        ' tail. The "-0" is a marker, checked below, so this
                        ' subtraction is only appended once per cell even if
                        ' Input_Stuff runs again on top of an already-patched sheet.
                        If Not (Right(.Cells(ivalue + 1, j).Formula, 2) = "-0") Then
                            .Cells(ivalue + 1, j).Formula = .Cells(ivalue + 1, j).Formula & "-" & charCheck(j + 64) & (dval + 3) & "-" & charCheck(j + 64) & (dval + 5) & "-0"
                        End If
                    End If
                End With

                ' ---------------------------------------------------------------
                ' DOUBLED LINE CHECK (earlier in sheet)
                ' Scan all earlier rows for an identical line (same attributes,
                ' same mark code). If found, subtract those counts too.
                ' This handles the case where the same variety/size appears
                ' multiple times in the Pakplan (e.g. split across two batches).
                ' ---------------------------------------------------------------
                Dim dubdub As Boolean
                dubdub = False
                For k = 1 To (vorRows - 5)
                    With ThisWorkbook.Worksheets(shName)
                        If ((k < vals) And _
                            (.Cells(vals, 3).Value = .Cells(k, 3).Value) And _
                            (.Cells(vals, 5).Value = .Cells(k, 5).Value) And _
                            (.Cells(vals, 6).Value = .Cells(k, 6).Value) And _
                            (.Cells(vals, 7).Value = .Cells(k, 7).Value) And _
                            (.Cells(vals, 8).Value = .Cells(k, 8).Value) And _
                            (.Cells(vals, 9).Value = .Cells(k, 9).Value) And _
                            (.Cells(vals, 10).Value = .Cells(k, 10).Value) And _
                            (.Cells(vals, 11).Value = .Cells(k, 11).Value) And _
                            (.Cells(vals, 16).Value = .Cells(k, 16).Value) And _
                            ((Not IsEmpty(.Cells(vals, j))) And (Not IsEmpty(.Cells(k, j))))) Then
                            ' Subtract the earlier row's Stock and Dispatched counts
                            If sPack = "PACKHOUSE" Then
                                sForm = sForm & "-" & charCheck(j + 64) & (k + 3) & "-" & charCheck(j + 64) & (k + 5)
                                doubledCheck = doubledCheck + 1
                            Else
                                sForm = sForm & "-" & charCheck(j + 64) & (k + 5)
                                doubledCheck = doubledCheck + 1
                            End If
                            dubdub = True
                        End If

                        ' Also check for the "D-prefix" neighbour pattern in earlier rows
                        ivalue = k
                        dval = k + 6
                        If (k > 3) And (dval < (vorRows - 5)) And _
                           (.Cells(dval, 3).Value = .Cells(ivalue, 3).Value) And _
                           (.Cells(dval, 5).Value = .Cells(ivalue, 5).Value) And _
                           (.Cells(dval, 6).Value = .Cells(ivalue, 6).Value) And _
                           (.Cells(dval, 7).Value = .Cells(ivalue, 7).Value) And _
                           (.Cells(dval, 8).Value = .Cells(ivalue, 8).Value) And _
                           (.Cells(dval, 9).Value = .Cells(ivalue, 9).Value) And _
                           (.Cells(dval, 10).Value = .Cells(ivalue, 10).Value) And _
                           (.Cells(dval, 11).Value = .Cells(ivalue, 11).Value) And _
                           (.Cells(dval, 16).Value = "D" & .Cells(ivalue, 16).Value) And _
                           (Not IsEmpty(.Cells(ivalue, 3))) Then
                            If Not (Right(.Cells(ivalue + 1, j).Formula, 2) = "-0") Then
                                ' FINDING: this is a leftover debug prompt, not a
                                ' user-facing warning - it just dumps the cell's
                                ' current value with no explanation of what it means
                                ' or why it's being shown. If this condition is ever
                                ' true during a real Update Data run, whoever is
                                ' sitting at the keyboard gets a blocking MsgBox they
                                ' can't make sense of. Worth deciding whether to
                                ' remove it, or turn it into a real message/log entry.
                                MsgBox (.Cells(ivalue + 1, j).Value)
                            End If
                        End If

                        ' Once we reach the current row, stop scanning
                        If (k = vals) Then k = vorRows - 5
                    End With
                Next k

                ' Write the completed formula to the cell
                ThisWorkbook.Worksheets(shName).Cells(i, j).Formula = sForm

            Next j   ' Next size column
        End If   ' vals <> 0
    Next i   ' Next Vordering row

    ' --- Activate Vordering and force a recalc ---
    ThisWorkbook.Worksheets(shName).Select
    OptimizeVBA (False)
    DoEvents
    OptimizeVBA (True)
    curper = totPerc

    ' =========================================================================
    ' SECOND PASS: Fix negative Outstanding values caused by doubled lines.
    ' When doubledCheck > 0, some Outstanding rows may show negative values
    ' because a duplicate line was counted twice and then subtracted.
    ' This pass identifies those rows and redistributes Stock/Dispatched counts
    ' between them until Outstanding = 0 (or as close as possible).
    ' =========================================================================
    If doubledCheck > 0 Then
        For i = 1 To (vorRows - 5)

            totPerc = curper + Round((i / (vorRows - 5)) * 7, 0)
            If totPerc > 100 Then totPerc = 100
            UserForm1.Label1.Caption = Str(totPerc) + "% Completed"
            UserForm1.Label2.Width = totPerc * 2
            UserForm1.Label3.Caption = "Finalizing..."
            DoEvents

            Dim isize As Integer

            ' Find rows where Outstanding is negative (over-counted)
            If (ThisWorkbook.Worksheets(shName).Cells(i, 12).Value = "Pallets Outstanding" And _
                ThisWorkbook.Worksheets(shName).Cells(i, 16).Value < 0) Then

                Dim stillLeft As Integer
                stillLeft = 0
                Dim arrNextCount() As Integer
                Dim iNumb As Integer
                iNumb = i - 2
                ' Find the next occurrence of the same line (the duplicate).
                ' NOTE: arrNextCount is only ever filled when stillLeft first
                ' becomes 1 (the "If stillLeft = 1 Then" guard below) - it snapshots
                ' whichever matching row is found FIRST as k counts up. If a third
                ' or later matching row also exists further down the sheet,
                ' stillLeft keeps incrementing for it, but its per-size-column
                ' counts are never captured into arrNextCount. In practice this
                ' only matters when the same attributes appear 3+ times, which the
                ' redistribution logic below doesn't otherwise account for either.
                For k = i To (vorRows - 5)
                    With ThisWorkbook.Worksheets(shName)
                        If ((.Cells(iNumb, 3).Value = .Cells(k, 3).Value) And _
                            (.Cells(iNumb, 5).Value = .Cells(k, 5).Value) And _
                            (.Cells(iNumb, 6).Value = .Cells(k, 6).Value) And _
                            (.Cells(iNumb, 7).Value = .Cells(k, 7).Value) And _
                            (.Cells(iNumb, 8).Value = .Cells(k, 8).Value) And _
                            (.Cells(iNumb, 9).Value = .Cells(k, 9).Value) And _
                            (.Cells(iNumb, 10).Value = .Cells(k, 10).Value) And _
                            (.Cells(iNumb, 11).Value = .Cells(k, 11).Value) And _
                            (.Cells(iNumb, 16).Value = .Cells(k, 16).Value) And _
                            (Not IsEmpty(.Cells(iNumb, j))) And (Not IsEmpty(.Cells(k, j)))) Then
                            stillLeft = stillLeft + 1
                            isize = totCol - 17
                            ReDim arrNextCount(isize)
                            If stillLeft = 1 Then
                                For j = 17 To totCol - 1
                                    If Not IsEmpty(.Cells(k, j)) Then
                                        If .Cells(k, j) = "*" Then
                                            arrNextCount(j - 17) = 10000  ' "*" means unlimited
                                        Else
                                            arrNextCount(j - 17) = .Cells(k, j).Value
                                        End If
                                    Else
                                        arrNextCount(j - 17) = 0
                                    End If
                                Next j
                            End If
                        End If
                    End With

                If stillLeft > 0 Then
                    Dim arrDisp() As Integer, arrStock() As Integer
                    Dim changable As Integer
                    changable = 0
                    Dim iNeed As Integer, iDisp As Integer, iOut As Integer, iStock As Integer
                    isize = totCol - 17
                    ReDim arrDisp(isize)
                    ReDim arrStock(isize)
                    iNeed = ThisWorkbook.Worksheets(shName).Cells(i - 1, 16).Value
                    iOut = ThisWorkbook.Worksheets(shName).Cells(i, 16).Value
                    iStock = ThisWorkbook.Worksheets(shName).Cells(i + 1, 16).Value
                    iDisp = ThisWorkbook.Worksheets(shName).Cells(i + 3, 16).Value

                    ' `changable` is the running total, across size columns, of how
                    ' many pallets are still available to move for this phase - it's
                    ' only counted from size columns where arrNextCount(j-17) > 0,
                    ' i.e. columns where the duplicate row actually has stock to draw
                    ' from. Each Do-loop iteration below moves exactly one pallet in
                    ' one column and decrements `changable`, so the loop is
                    ' guaranteed to terminate once every movable pallet has been used
                    ' or the target (iDisp = iNeed / iOut = 0) is reached.

                    ' --- Phase 1: Reduce Dispatched to match Needed ---
                    ' If more was dispatched than needed, move some back to Stock
                    If ThisWorkbook.Worksheets(shName).Cells(i + 3, 16).Value > ThisWorkbook.Worksheets(shName).Cells(i - 1, 16).Value Then
                        For j = 17 To totCol - 1
                            arrStock(j - 17) = ThisWorkbook.Worksheets(shName).Cells(i + 1, j).Value
                            arrDisp(j - 17) = ThisWorkbook.Worksheets(shName).Cells(i + 3, j).Value
                            If arrNextCount(j - 17) > 0 Then
                                changable = changable + ThisWorkbook.Worksheets(shName).Cells(i + 3, j).Value
                            End If
                        Next j
                        Do Until (iDisp = iNeed) Or (changable = 0)
                            For j = 17 To totCol - 1
                                If arrNextCount(j - 17) > 0 Then
                                    If iDisp > iNeed Then
                                        If arrDisp(j - 17) > ThisWorkbook.Worksheets(shName).Cells(i - 1, j).Value Then
                                            arrDisp(j - 17) = arrDisp(j - 17) - 1
                                            arrStock(j - 17) = arrStock(j - 17) + 1
                                            ThisWorkbook.Worksheets(shName).Cells(i + 3, j).Value = ThisWorkbook.Worksheets(shName).Cells(i + 3, j).Value - 1
                                            iDisp = iDisp - 1
                                            changable = changable - 1
                                        End If
                                    End If
                                End If
                            Next j
                        Loop
                        ' Clear stock row values that were moved to dispatched
                        For j = 17 To totCol - 1
                            If arrStock(j - 17) > 0 Then
                                arrStock(j - 17) = 0
                                ThisWorkbook.Worksheets(shName).Cells(i + 1, j).Value = 0
                            End If
                        Next j
                        OptimizeVBA (False)
                        DoEvents
                        OptimizeVBA (True)
                    End If

                    ' --- Phase 2: Reduce Stock to resolve remaining negative Outstanding ---
                    changable = 0
                    iNeed = ThisWorkbook.Worksheets(shName).Cells(i - 1, 16).Value
                    iOut = ThisWorkbook.Worksheets(shName).Cells(i, 16).Value
                    iStock = ThisWorkbook.Worksheets(shName).Cells(i + 1, 16).Value
                    iDisp = ThisWorkbook.Worksheets(shName).Cells(i + 3, 16).Value
                    For j = 17 To totCol - 1
                        arrStock(j - 17) = ThisWorkbook.Worksheets(shName).Cells(i + 1, j).Value
                        arrDisp(j - 17) = ThisWorkbook.Worksheets(shName).Cells(i + 3, j).Value
                        If arrNextCount(j - 17) > 0 Then
                            changable = changable + ThisWorkbook.Worksheets(shName).Cells(i + 1, j).Value
                        End If
                    Next j
                    If (iStock + iDisp) > iNeed Then
                        Do Until (iOut = 0) Or (changable = 0)
                            For j = 17 To totCol - 1
                                If arrNextCount(j - 17) > 0 Then
                                    If arrStock(j - 17) > 0 Then
                                        ThisWorkbook.Worksheets(shName).Cells(i + 1, j).Value = ThisWorkbook.Worksheets(shName).Cells(i + 1, j).Value - 1
                                        arrStock(j - 17) = arrStock(j - 17) - 1
                                        iStock = iStock - 1
                                        iOut = iOut + 1
                                        changable = changable - 1
                                    End If
                                End If
                                If iOut = 0 Then j = totCol - 1   ' Early exit once resolved
                            Next j
                        Loop
                        OptimizeVBA (False)
                        DoEvents
                        OptimizeVBA (True)
                    End If
                End If   ' stillLeft > 0
                
                Next k
            End If   ' Pallets Outstanding < 0
        Next i
    End If   ' doubledCheck > 0

    OptimizeVBA (False)
    If ribref Then UserForm1.Hide

    Exit Sub
ErrHandler:
    HandleModuleError "Input_Stuff"
End Sub


' =============================================================================
' COPY_STUFF
' Copies the entire Vordering sheet range (A1 through to 2 columns past TOTAL)
' to the clipboard. Used as a quick manual clipboard helper.
' =============================================================================
Public Sub Copy_Stuff()
    Dim finalcol As Integer
    Dim finRow As Integer
    finalcol = ThisWorkbook.Worksheets("Vordering").Range("Q" & startline & ":AZ" & startline).Find("TOTAL", , xlValues, xlWhole).Column + 2
    finRow = ThisWorkbook.Worksheets("Vordering").Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    ThisWorkbook.Worksheets("Vordering").Range("A1:" & charCheck(finalcol + 64) & finRow).Copy
End Sub


' =============================================================================
' RIBBON WRAPPER: Export_Stuff_R
' =============================================================================
Public Sub Export_Stuff_R(control As IRibbonControl)
    Export_Stuff
End Sub


' =============================================================================
' EXPORT_STUFF
' Saves a clean export copy of the workbook (values only for Vordering and
' Opsomming; full copy for Pakplan and Grafieke), then emails it.
'
' WHAT IT DOES:
'   1. Asks the user to confirm Save & Send.
'   2. Creates a new workbook.
'   3. Copies Pakplan sheet fully (with formulas and formatting).
'   4. Copies Vordering as values+formats only (no live formulas).
'      Deletes the top rows above startline-3 (internal control rows).
'   5. Copies Opsomming as values+formats only.
'   6. Copies Grafieke sheet with charts intact.
'   7. Saves the new workbook to the path defined by mSave + variety + filename.
'   8. Calls Mail_Stuff to send it via Outlook (unless on the server PC).
'
' FILE NAMING:
'   [mSave]\[variety]\[tName] Pakplan Vordering [WeekNumb].xlsx
'   e.g. C:\Reports\HVN\Havalina Pakplan Vordering Week 7.xlsx
'
' NOTE: If run on "PALTRACK-PC" (the server), no email is sent because
'       Outlook is not set up there. A reminder message is shown instead.
' =============================================================================
Public Sub Export_Stuff()
    On Error GoTo ErrHandler
    gAbortPipeline = False
    If Not EnsureEntitled() Then Exit Sub
    InitiateConstants

    Dim answ As String
    answ = MsgBox("Do you want to Save & Send?", vbQuestion + vbYesNo, "User Response")

    If answ = "6" Then   ' User confirmed
        Dim wbSource As Workbook, wbNew As Workbook
        Dim wsCopy As Worksheet, wsValues As Worksheet
        Dim wsSummary As Worksheet, wsChart As Worksheet, wsNew As Worksheet
        Dim tlen As Integer, i As Integer
        Dim rng As Range, col As Range
        Dim SavePath As String
        Dim tempN As String

        Application.ScreenUpdating = False
        vari = ThisWorkbook.Worksheets("Vordering").Cells(4, 4).Value

        Set wbSource = ThisWorkbook
        Set wsCopy = wbSource.Worksheets("Pakplan")
        Set wsValues = wbSource.Worksheets("Vordering")
        Set wsSummary = wbSource.Worksheets("Opsomming")
        Set wsChart = wbSource.Worksheets("Grafieke")

        ' --- Parse tName: extract the pack type name from the Pakplan title cell ---
        ' The title cell contains something like "Havalina Pakplan Vordering Week 7"
        ' We want just "Havalina" (everything before the word "PACK")
        tName = ""
        tempN = wsValues.Range("A" & (startline - 2)).Value
        tlen = Len(tempN)
        For i = 1 To tlen
            tName = Mid(tempN, i, 4)
            If UCase(tName) = "PACK" Then
                tName = Left(tempN, i - 2)
                i = tlen   ' Exit loop
            End If
        Next i

        ' Sanitise tName by replacing invalid filename characters
        Dim invalidChars As Variant, chch As Variant
        invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        For Each chch In invalidChars
            tName = Replace(tName, chch, " ")
        Next chch
        tName = StrConv(tName, vbProperCase)

        ' --- Parse WeekNumb: extract the week number from the title cell ---
        ' Looks for the word "WEEK" and takes the 2 characters after "WEEK "
        WeekNumb = ""
        tempN = wsValues.Range("A" & (startline - 2)).Value
        tlen = Len(tempN)
        For i = 1 To tlen
            WeekNumb = Mid(tempN, i, 4)
            If UCase(WeekNumb) = "WEEK" Then
                WeekNumb = Mid(tempN, i + 5, 2)
                i = tlen
            End If
        Next i
        If Not (IsNumeric(WeekNumb)) Then WeekNumb = Left(WeekNumb, 1)
        weeknumber = WeekNumb
        WeekNumb = "Week " & WeekNumb

        ' Build the save path
        SavePath = mSave & vari & "\" & tName & " Pakplan Vordering " & WeekNumb & ".xlsx"

        ' --- Create new export workbook ---
        Set wbNew = Workbooks.Add

        ' Copy Pakplan sheet (full, with formulas)
        wsCopy.Copy Before:=wbNew.Sheets(1)
        wbNew.Sheets(1).name = wsCopy.name

        ' Copy Vordering as values + formats (no live formulas in exported file)
        Set wsNew = wbNew.Sheets.Add(After:=wbNew.Sheets(wbNew.Sheets.Count))
        wsNew.name = wsValues.name
        wsValues.Cells.Copy
        wsNew.Cells.PasteSpecial Paste:=xlPasteValues
        wsNew.Cells.PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
        For Each col In wsValues.UsedRange.Columns
            wsNew.Columns(col.Column).ColumnWidth = wsValues.Columns(col.Column).ColumnWidth
        Next col
        ' Remove the internal control rows (above the visible data header)
        If startline - 3 > 0 Then wsNew.Rows("1:" & (startline - 3)).Delete
        wsNew.Range("A1").Select

        ' Copy Opsomming as values + formats
        Set wsNew = wbNew.Sheets.Add(After:=wbNew.Sheets(wbNew.Sheets.Count))
        wsNew.name = wsSummary.name
        wsSummary.Cells.Copy
        wsNew.Cells.PasteSpecial Paste:=xlPasteValues
        wsNew.Cells.PasteSpecial Paste:=xlPasteFormats
        Application.CutCopyMode = False
        For Each col In wsSummary.UsedRange.Columns
            wsNew.Columns(col.Column).ColumnWidth = wsSummary.Columns(col.Column).ColumnWidth
        Next col
        wsNew.Range("A1").Select

        ' Copy Grafieke (charts) - full copy to preserve chart objects
        wsChart.Copy After:=wbNew.Sheets(wbNew.Sheets.Count)
        wbNew.Sheets(wbNew.Sheets.Count).name = wsChart.name

        ' Set view and delete the default empty Sheet1
        wbNew.Sheets("Vordering").Select
        wbNew.Sheets("Vordering").Range("A1").Select
        ActiveWindow.Zoom = 80
        Application.DisplayAlerts = False
        wbNew.Worksheets("Sheet1").Delete
        If Dir(SavePath) <> "" Then Kill SavePath  ' Delete if file already exists
        wbNew.SaveAs Filename:=SavePath, FileFormat:=xlOpenXMLWorkbook
        Application.DisplayAlerts = True
        wbNew.Close SaveChanges:=False

        Application.ScreenUpdating = True

        ' Notify user and send email (skip email on server PC)
        If Environ("COMPUTERNAME") = "PALTRACK-PC" Then
            MsgBox ("No email is set up on the server, please copy saved file and download to device with a valid email to send from.")
            MsgBox "Workbook saved as " & SavePath, vbInformation, "Export Complete"
        Else
            MsgBox "Workbook saved as " & SavePath, vbInformation, "Export Complete"
            Mail_Stuff SavePath, weeknumber
        End If
    End If

    Exit Sub
ErrHandler:
    HandleModuleError "Export_Stuff"
End Sub


' =============================================================================
' MAIL_STUFF
' Creates and displays a pre-populated Outlook email with the exported
' workbook as an attachment.
'
' PARAMETERS:
'   sPath - Full path to the saved export file
'   wkn   - Week number (Integer) used in the email subject and body
'
' NOTE: The email is DISPLAYED (not auto-sent) so the user can review it
'       before sending. This is intentional for quality control.
'
' The greeting text differs between "Ohr" (Afrikaans, no "Groete" sign-off)
' and "Junction" (Afrikaans, includes "Groete"). This is determined by
' checking mBCC against a known email address.
' TODO: Consider a more robust way to distinguish these contexts (maybe
'       a named cell or constant instead of checking an email address).
' =============================================================================
Sub Mail_Stuff(sPath As String, wkn As Integer)
    Dim OutApp As Object
    Dim OutMail As Object
    Dim sVar As String
    Dim pos As Integer
    pos = InStrRev(sPath, "\")
    sVar = Mid(sPath, pos - 3, 3)

    ' Record the send timestamp in the Data sheet ribbon label cell
    ThisWorkbook.Worksheets("Data").Select
    OptimizeVBA (False)
    SetAppSetting "LastSent", "Last Sent: " & Now()
    DoEvents
    ForceRibbonRefresh
    OptimizeVBA (True)

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)   ' 0 = olMailItem

    On Error Resume Next
    With OutMail
        .To = mTo
        .CC = mCC
        .BCC = mBCC
        .Subject = sVar & " Pakplan Vordering Week " & wkn
        .Display   ' Show the email for review before sending
        ' Set email body based on context (Ohr vs Junction)
        If mContext = "Ohr" Then
            ' Ohr context: no sign-off line
            .HTMLBody = "<font style=""font-family: Aptos; font-size: 13pt;"">Goeiedag,<br><br>Sien aangeheg die Pakplan Vordering vir Week " & wkn & ".</font>" & .HTMLBody
        Else
            ' Junction context: include "Groete" sign-off
            .HTMLBody = "<font style=""font-family: Aptos; font-size: 13pt;"">Goeiedag,<br><br>Sien aangeheg die Pakplan Vordering vir Week " & wkn & ".<br><br>Groete</font>" & .HTMLBody
        End If
        .Attachments.Add sPath
    End With
    On Error GoTo 0

    ' Re-enable events and screen updating
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    Set OutMail = Nothing
    Set OutApp = Nothing

    ThisWorkbook.Worksheets("Vordering").Select
    OptimizeVBA (False)
End Sub


' =============================================================================
' RIBBON WRAPPER: Short_Stuff_R
' =============================================================================
Public Sub Short_Stuff_R(control As IRibbonControl)
    Short_Stuff
End Sub


' =============================================================================
' SHORT_STUFF
' Builds or rebuilds the "Opsomming" (Summary) sheet.
'
' WHAT IT DOES:
' Reads the Vordering sheet and extracts one row per packing line into
' Opsomming, showing:
'   - Columns A-P: copied directly from the Vordering data row
'   - Columns Q onwards (size columns): shows Stock + Dispatched totals
'     (i.e. how many were actually packed, to compare against the plan)
'   - Final column (STILL NEED): the Outstanding value from column P
'
' Rows where Outstanding = 0 are highlighted green.
' Also applies borders and alignment.
'
' NOTE: For columns >= 81 (past column Q in ASCII), the value shown is
'       Stock (i+1) + Dispatched (i+3) rather than the "Needed" row,
'       because that reflects actual production better for reporting.
' =============================================================================
Sub Short_Stuff()
    On Error GoTo ErrHandler
    gAbortPipeline = False
    If Not EnsureEntitled() Then Exit Sub

    shName = "Vordering"
    Dim short As String
    Dim sanswer As String
    Dim checkLoad As Boolean
    Dim prog As Double
    Dim iRows As Integer

    checkLoad = False
    short = "Opsomming"
    sanswer = "6"
    totCol = ThisWorkbook.Sheets(shName).Range("Q" & startline & ":AZ" & startline).Find("TOTAL", , xlValues, xlWhole).Column + 64
    iRows = (ThisWorkbook.Sheets(shName).Range("A" & startline & ":A" & 1000).Find("GRAND TOTA*", , xlValues, xlWhole).Row + 5) * 3

    ' Create the sheet if needed, otherwise just select it
    If Not sheetExists(short) Then
        ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).name = short
        sanswer = "6"
    Else
        ThisWorkbook.Sheets(short).Select
    End If

    ' Show progress bar if called standalone
    If (totPerc = 0) Or (IsEmpty(totPerc)) Then
        UserForm1.Show (False)
        UserForm1.Width = 220
        UserForm1.Frame1.Width = 200
        UserForm1.Height = 98
        UserForm1.StartUpPosition = 0
        UserForm1.Caption = "Progress Bar"
        UserForm1.Label1.Caption = "0% Completed"
        UserForm1.Label2.Caption = ""
        UserForm1.Label3.Caption = "Routing to Summary..."
        UserForm1.Label2.Width = 0
        UserForm1.Label2.Height = UserForm1.Frame1.Height - 4
        UserForm1.Label2.BackColor = vbRed
        UserForm1.Frame1.Caption = ""
        checkLoad = True
    End If

    If sanswer = "6" Then
        ' Clear old data from summary sheet
        ThisWorkbook.Worksheets(short).Rows(1 & ":" & Round(iRows / 5)).Delete
        OptimizeVBA (True)

        ' Copy the header row from Vordering to Opsomming row 1
        ThisWorkbook.Sheets(shName).Range("A" & startline & ":" & charCheck(totCol) & startline).Copy _
            Destination:=ThisWorkbook.Sheets(short).Range("A" & 1).End(xlUp)
        ' Copy column widths to match
        ThisWorkbook.Sheets(shName).Range("A" & startline & ":" & charCheck(totCol) & startline).Copy
        ThisWorkbook.Sheets(short).Range("A" & startline & ":" & charCheck(totCol) & startline).PasteSpecial xlPasteColumnWidths

        ' Style the header row with borders
        With ThisWorkbook.Sheets(short).Range("A" & 1 & ":" & charCheck(totCol) & 1).Borders()
            .LineStyle = xlContinuous
            .Color = vbBlack
            .Weight = xlMedium
        End With

        ' Add a "STILL NEED" column header in the last column
        ThisWorkbook.Sheets(short).Range(charCheck(totCol) & "1").Value = "STILL NEED"
        ThisWorkbook.Sheets(short).Range(charCheck(totCol) & "1").WrapText = True

        ' Find the last row of data in Vordering
        Dim FinCell As Range, finRow As Integer
        Set FinCell = ThisWorkbook.Worksheets(shName).Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
        finRow = FinCell.Row

        Dim i As Integer, j As Integer, m As Integer
        Dim k As Integer
        k = 1   ' Current row in Opsomming (starts at 1 = header)

        Dim sLine(16) As String

        ' --- Main loop: copy "Pallets Outstanding" rows from Vordering to Opsomming ---
        For i = startline To finRow

            ' Update progress bar
            prog = (i - startline) / (finRow - startline)
            UserForm1.Label2.BackColor = GetProgressColor(CDbl(prog))
            If checkLoad = True Then
                totPerc = Round((i / finRow) * 100, 0)
                If totPerc > 100 Then totPerc = 100
                UserForm1.Label1.Caption = Str(totPerc) + "% Completed"
                UserForm1.Label2.Width = totPerc * 2
                UserForm1.Label3.Caption = "Updating Summary..."
                DoEvents
            End If

            ' Only process "Pallets Outstanding" rows
            If ThisWorkbook.Sheets(shName).Range("L" & i).Value = "Pallets Outstanding" Then
                k = k + 1

                ' Copy all columns for this row
                For j = 65 To (totCol - 1)
                    If j < 81 Then
                        ' Columns A-P (ASCII 65-80): copy directly from the data row (i-2)
                        ThisWorkbook.Sheets(short).Range(charCheck(j) & k).Value = ThisWorkbook.Sheets(shName).Range(charCheck(j) & (i - 2)).Value
                    Else
                        ' Size columns (Q onwards): show Stock + Dispatched (packed so far)
                        If ThisWorkbook.Sheets(shName).Range(charCheck(j) & (i + 1)).Value + ThisWorkbook.Sheets(shName).Range(charCheck(j) & (i + 3)).Value > 0 Then
                            ThisWorkbook.Sheets(short).Range(charCheck(j) & k).Value = _
                                ThisWorkbook.Sheets(shName).Range(charCheck(j) & (i + 1)).Value + _
                                ThisWorkbook.Sheets(shName).Range(charCheck(j) & (i + 3)).Value
                        Else
                            ThisWorkbook.Sheets(short).Range(charCheck(j) & k).Formula = ""
                        End If
                    End If
                Next j

                ' Apply borders to this summary row
                With ThisWorkbook.Sheets(short).Range("A" & k & ":" & charCheck(totCol) & k).Borders()
                    .LineStyle = xlContinuous
                    .Color = vbBlack
                    .Weight = xlThin
                End With
                ' Medium border on the right of col P (separates info from size columns)
                With ThisWorkbook.Sheets(short).Range("P" & k & ":P" & k).Borders(xlRight)
                    .LineStyle = xlContinuous
                    .Color = vbBlack
                    .Weight = xlMedium
                End With
                ' Medium borders around the STILL NEED column
                With ThisWorkbook.Sheets(short).Range(charCheck(totCol) & k).Borders(xlLeft)
                    .LineStyle = xlContinuous
                    .Color = vbBlack
                    .Weight = xlMedium
                End With
                With ThisWorkbook.Sheets(short).Range(charCheck(totCol) & k).Borders(xlRight)
                    .LineStyle = xlContinuous
                    .Color = vbBlack
                    .Weight = xlMedium
                End With

                ' Write Outstanding value in the STILL NEED column
                ThisWorkbook.Sheets(short).Range(charCheck(totCol) & k).Value = ThisWorkbook.Sheets(shName).Range("P" & i).Value

                ' Highlight green if nothing is outstanding
                If ThisWorkbook.Sheets(short).Range(charCheck(totCol) & k).Value = 0 Then
                    ThisWorkbook.Sheets(short).Range(charCheck(totCol) & k).Interior.Color = RGB(70, 170, 100)
                End If

            End If
        Next i

        ' Centre-align all size columns in Opsomming
        ThisWorkbook.Sheets(short).Range("Q1:" & charCheck(totCol) & k).HorizontalAlignment = xlCenter

        ' Medium top border on the last data row and the row after (visual separator)
        With ThisWorkbook.Sheets(short).Range("A" & k & ":" & charCheck(totCol) & k).Borders(xlTop)
            .LineStyle = xlContinuous
            .Color = vbBlack
            .Weight = xlMedium
        End With
        With ThisWorkbook.Sheets(short).Range("A" & k + 1 & ":" & charCheck(totCol) & k + 1).Borders(xlTop)
            .LineStyle = xlContinuous
            .Color = vbBlack
            .Weight = xlMedium
        End With

    End If  ' sanswer = "6"

    OptimizeVBA (False)
    ThisWorkbook.Worksheets(short).Select
    If checkLoad = True Then UserForm1.Hide

    ' Clear clipboard and return to top of sheet
    ThisWorkbook.Sheets(short).Range("AZ1").Copy
    ThisWorkbook.Worksheets(short).Range("A1").Select
    totPerc = 0

    Exit Sub
ErrHandler:
    HandleModuleError "Short_Stuff"
End Sub


' =============================================================================
' RIBBON WRAPPER: Chart_Stuff_R
' =============================================================================
Public Sub Chart_Stuff_R(control As IRibbonControl)
    Chart_Stuff
End Sub


' =============================================================================
' CHART_STUFF
' Builds or rebuilds the "Grafieke" (Charts) sheet.
'
' WHAT IT DOES:
' For each packing line in Vordering, creates a small pie chart showing:
'   - Green slice:  Pallets Packed (Stock + Dispatched)
'   - Yellow slice: Pallets Outstanding
'   - Red slice:    Pallets Overpacked
'
' Charts are laid out in a grid (6 columns wide) with each chart 180x140 px.
' The chart title includes: variety - pack - grade - batch number (and mark if present).
'
' Data is read from the column identified as "BATC*" (the Batch NR column)
' since that column's row offsets conveniently map to each of the 5 sub-rows
' via the 6-row block structure.
'
' NOTE: Items with zero pallets needed are shown as an empty chart with "0"
'       label, to maintain the grid position for all items.
' =============================================================================
Sub Chart_Stuff()
    On Error GoTo ErrHandler
    gAbortPipeline = False
    If Not EnsureEntitled() Then Exit Sub

    shName = "Vordering"
    Dim sanswer As String
    Dim sChart As String
    Dim iItems As Integer
    Dim i As Integer
    Dim checkLoad As Boolean
    Dim prog As Double
    Dim needed As Double
    Dim outstanding As Double
    Dim inStock As Double
    Dim dispatched As Double
    Dim overpacked As Double
    Dim batCol As Integer
    Dim datArray As Variant
    Dim namArray As Variant
    Dim topOffset As Double
    Dim leftOffset As Double
    Dim rowMax As Long: rowMax = 6   ' Number of chart columns in the grid
    Dim chartObj As ChartObject
    checkLoad = False

    ' Locate the BATCH NR column in Vordering (the reference column for reading row offsets)
    batCol = ThisWorkbook.Sheets(shName).Range("A" & startline & ":AZ" & startline).Find("BATC*", , xlValues, xlWhole).Column + 64

    ' Count the number of packing lines (items) = rows between startline and GRAND TOTAL, divided by block size (6)
    iItems = (ThisWorkbook.Sheets(shName).Range("A" & startline & ":A" & 1000).Find("GRAND TOTA*", , xlValues, xlWhole).Row - startline - 1) / 6

    sChart = "Grafieke"
    OptimizeVBA (True)

    ' Create the charts sheet if it doesn't exist
    If Not sheetExists(sChart) Then
        ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)).name = sChart
    Else
        ThisWorkbook.Sheets(sChart).Select
    End If

    ' Delete all existing charts on the sheet before rebuilding
    For Each chartObj In ThisWorkbook.Sheets(sChart).ChartObjects
        chartObj.Delete
    Next

    ' Show progress bar if called standalone
    If (totPerc = 0) Or (IsEmpty(totPerc)) Then
        UserForm1.Show (False)
        UserForm1.Width = 220
        UserForm1.Frame1.Width = 200
        UserForm1.Height = 98
        UserForm1.StartUpPosition = 0
        UserForm1.Caption = "Progress Bar"
        UserForm1.Label1.Caption = "0% Completed"
        UserForm1.Label2.Caption = ""
        UserForm1.Label3.Caption = "Routing to Charts..."
        UserForm1.Label2.Width = 0
        UserForm1.Label2.Height = UserForm1.Frame1.Height - 4
        UserForm1.Label2.BackColor = vbRed
        UserForm1.Frame1.Caption = ""
        checkLoad = True
    End If

    ' --- Main loop: create one pie chart per packing line ---
    For i = 1 To iItems

        ' Update progress bar
        prog = (i - 1) / (iItems - 1)
        UserForm1.Label2.BackColor = GetProgressColor(CDbl(prog))
        If checkLoad = True Then
            totPerc = Round((i / iItems) * 100, 0)
            If totPerc > 100 Then totPerc = 100
            UserForm1.Label1.Caption = Str(totPerc) + "% Completed"
            UserForm1.Label2.Width = totPerc * 2
            UserForm1.Label3.Caption = "Updating Charts..."
            DoEvents
        End If

        ' Read pallet counts from the Vordering block for this item.
        ' Row offset formula: 6*i - startline + sub-row offset
        '   +3 = Pallets Outstanding row
        '   +4 = Pallets in Stock row
        '   +5 = Pallets Overpacked row
        '   +6 = Pallets Dispatched row
        With ThisWorkbook.Worksheets(shName)
            needed = .Cells(6 * i - startline + 2, charCheck(batCol)).Value
            outstanding = .Cells(6 * i - startline + 3, charCheck(batCol)).Value
            If outstanding < 0 Then outstanding = 0   ' Don't show negative outstanding
            inStock = .Cells(6 * i - startline + 4, charCheck(batCol)).Value
            dispatched = .Cells(6 * i - startline + 6, charCheck(batCol)).Value
            overpacked = .Cells(6 * i - startline + 5, charCheck(batCol)).Value
            If (overpacked > 0) And (inStock > needed) Then inStock = needed
        End With

        ' Chart data: Packed (in stock + dispatched), Outstanding, Overpacked
        datArray = Array(inStock + dispatched, outstanding, overpacked)
        namArray = Array("Packed", "Outstanding", "Overpack")

        ' Calculate position in the grid (6 charts per row)
        leftOffset = ((i - 1) Mod rowMax) * 185
        topOffset = Int((i - 1) / rowMax) * 145

        ' Create the chart object
        Set chartObj = ThisWorkbook.Sheets(sChart).ChartObjects.Add( _
            Left:=leftOffset, Top:=topOffset, Width:=180, Height:=140)

        ' Configure the pie chart
        With chartObj.Chart
            .ChartType = xlPie
            .SeriesCollection.NewSeries
            With .SeriesCollection(1)
                .ApplyDataLabels Type:=xlDataLabelsShowLabel
                .HasDataLabels = True
                ' If nothing was packed or needed, show a zero chart
                If inStock + dispatched + outstanding = 0 Then
                    .Values = Array(1, 0, 0)
                    .Points(1).DataLabel.text = "0"
                Else
                    .Values = datArray
                End If
                .XValues = namArray
                ' Slice colours: Green=Packed, Yellow=Outstanding, Red=Overpacked
                .Points(1).Format.Fill.ForeColor.RGB = RGB(0, 170, 0)
                .Points(2).Format.Fill.ForeColor.RGB = RGB(220, 220, 0)
                .Points(3).Format.Fill.ForeColor.RGB = RGB(220, 0, 0)
                .DataLabels.ShowValue = True
                .DataLabels.ShowCategoryName = False
                .DataLabels.Font.Bold = True
                .DataLabels.Font.Size = 11
                ' Remove data labels for zero slices (keeps chart clean)
                Dim pt As Point
                For Each pt In .Points
                    If pt.DataLabel.text = "0" Then pt.DataLabel.Delete
                Next pt
            End With
            ' Chart title: variety - pack - grade - batch (and mark if present)
            .HasTitle = True
            .ChartTitle.Format.TextFrame2.TextRange.Font.Size = 10
            .ChartTitle.text = _
                ThisWorkbook.Worksheets(shName).Cells(6 * i - startline + 1, charCheck(batCol - 13)).Value & " - " & _
                ThisWorkbook.Worksheets(shName).Cells(6 * i - startline + 1, charCheck(batCol - 11)).Value & " - Grade " & _
                ThisWorkbook.Worksheets(shName).Cells(6 * i - startline + 1, charCheck(batCol - 10)).Value & " - " & _
                ThisWorkbook.Worksheets(shName).Cells(6 * i - startline + 1, charCheck(batCol)).Value
            ' Append the mark code if it's present
            If ThisWorkbook.Worksheets(shName).Cells(6 * i - startline + 1, charCheck(batCol - 7)).Value <> "" Then
                .ChartTitle.text = .ChartTitle.text & " - " & _
                    ThisWorkbook.Worksheets(shName).Cells(6 * i - startline + 1, charCheck(batCol - 7)).Value
            End If
            .HasLegend = True
            .Legend.Font.Size = 9
        End With

    Next i

    OptimizeVBA (False)
    If checkLoad = True Then UserForm1.Hide
    totPerc = 0

    Exit Sub
ErrHandler:
    HandleModuleError "Chart_Stuff"
End Sub


' =============================================================================
' OPTIMIZE_VBA
' Toggles Excel performance optimisations on or off.
' Call OptimizeVBA(True) before a long operation and OptimizeVBA(False) after.
'
' WHAT IT DOES:
'   isOn = True:  Manual calculation + events/screen updating OFF (faster)
'   isOn = False: Automatic calculation + events/screen updating ON (normal)
'
' IMPORTANT: Always call OptimizeVBA(False) in error handlers or at the end
'            of every sub that calls OptimizeVBA(True), otherwise Excel will
'            be left in a non-responsive state.
' =============================================================================
Public Sub OptimizeVBA(isOn As Boolean)
    Application.Calculation = IIf(isOn, xlCalculationManual, xlCalculationAutomatic)
    Application.EnableEvents = Not (isOn)
    Application.ScreenUpdating = Not (isOn)
End Sub

' =============================================================================
' HANDLE MODULE ERROR
' Central error handler for the six main entry points (Setup/Update/Input/
' Short/Chart/Export_Stuff). Each has "On Error GoTo ErrHandler" at the top
' and calls this from its ErrHandler: label when something goes wrong.
'
' Restores Application state via OptimizeVBA(False) - see that Sub's own
' comment above for why this matters: without it, a crash mid-operation
' leaves Excel stuck in manual calculation with no screen updates, which
' looks like a second, unrelated bug on top of whatever actually failed.
' Then shows what failed and offers to email a report, so a failure
' doesn't just silently vanish into "it didn't work, don't know why".
' =============================================================================
Public Sub HandleModuleError(procName As String)
    Dim errNum As Long, errDesc As String
    errNum = Err.Number
    errDesc = Err.Description

    gAbortPipeline = True
    OptimizeVBA (False)
    On Error Resume Next
    UserForm1.Hide
    On Error GoTo 0

    Dim msg As String
    msg = procName & " ran into a problem and stopped:" & vbCrLf & vbCrLf & _
          "Error " & errNum & ": " & errDesc & vbCrLf & vbCrLf & _
          "Calculation and screen updating have been restored to normal." & vbCrLf & _
          "Send a report about this so it can be looked into?"

    If MsgBox(msg, vbYesNo + vbExclamation, "Pakplan Vordering - " & procName & " failed") = vbYes Then
        Dim context As String
        context = InputBox("Briefly describe what you were doing when this happened (optional):", "Error Report")
        SendErrorReport procName, errNum, errDesc, context
    End If
End Sub

