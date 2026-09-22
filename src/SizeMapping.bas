Attribute VB_Name = "SizeMapping"
Option Explicit

' =====================================================================
' SIZE MAPPING ENGINE  (replaces the old hand-permutated Select Case)
' =====================================================================
'
' HOW THIS WORKS:
'
'   - The actual carton pack for this line is read directly from
'     Cells(vals, 8).
'   - The Variety column (D:D) is NOT read - it can contain ambiguous
'     placeholder text like "Any seedless" or "Any variety" and it's
'     simply ignored. Instead, the family (VAL, STR, POM, or LEM) is
'     auto-detected from the numbers already present in the header
'     cell, by scoring how well each family's known counts explain
'     the WHOLE header (see DetectFamily below for why this has to
'     look at more than just the leading token).
'   - Once the family and its SIZE_REF are known, that pack's count
'     at that exact SIZE_REF is looked up directly from the
'     reference table.
'
'   - If more than one matching count is found (Z10D only - see
'     DetectFamily/GetSizeCriteria) -> sum them via chained COUNTIFS
'     blocks, same pattern as the original "20,22" special case,
'     generalized to any number of matches.
'   - If none appear -> fall back to a Large/Medium/Small marker
'     override, else the header's only token if there's just one,
'     else the first (outside-bracket) token as a last resort.
'
' TO ADD A NEW COUNT: add one line to SIZE_DATA. Nothing else needs
' to change - no new Case blocks, no new permutations, no variety
' lookups to update.
' =====================================================================

' Pipe-delimited reference data: VARIETY|SIZE_REF|PACK|COUNT
' NOTE: DEL, APV, MKN, GSV, CAR (Valencia-family cultivars), and EUR/SHD
' (which duplicated LEM/POM respectively) used to have their own blocks
' here, numerically identical to VAL/LEM/POM. Removed 2026-09 - no code
' path ever looks SIZE_DATA up by those varieties (family is detected
' purely from the header's numbers via FAMILY_ORDER, which only ever
' contains VAL/STR/POM/LEM - see DetectFamily below), so they were dead
' data. Of those varieties, the SEEDLESS cultivars are MKN, APV, and DEL;
' VAL, GSV, and CAR are not - recorded here since that fact has nowhere
' else to live now that the actual data rows are gone.
Private Const SIZE_DATA As String = _
    "LEM|0|A15C|56" & vbCrLf & "LEM|1|A15C|64" & vbCrLf & "LEM|2|A15C|75" & vbCrLf & "LEM|3|A15C|88" & vbCrLf & "LEM|3|A15C|100" & vbCrLf & "LEM|4|A15C|113" & vbCrLf & "LEM|5|A15C|138" & vbCrLf & "LEM|6|A15C|162" & vbCrLf & "LEM|6|A15C|189" & vbCrLf & "LEM|7|A15C|216" & vbCrLf & "LEM|1|E15D|70" & vbCrLf & "LEM|2|E15D|80" & vbCrLf & "LEM|3|E15D|90" & vbCrLf & "LEM|3|E15D|96" & vbCrLf & "LEM|1|E15C|70" & vbCrLf & "LEM|2|E15C|80" & vbCrLf & "LEM|3|E15C|90" & vbCrLf & "LEM|3|E15C|96" & vbCrLf & "LEM|1|D15D|70" & vbCrLf & "LEM|2|D15D|80" & vbCrLf & "LEM|3|D15D|90" & vbCrLf & "LEM|3|D15D|96" & vbCrLf & "LEM|1|C15D|70" & vbCrLf & "LEM|2|C15D|80" & vbCrLf & "LEM|3|C15D|90" & vbCrLf & "LEM|3|C15D|96" & vbCrLf & _
    "POM|2|A15C|6" & vbCrLf & "POM|3|A15C|8" & vbCrLf & "POM|3|A15C|10" & vbCrLf & "POM|4|A15C|12" & vbCrLf & "POM|5|A15C|14" & vbCrLf & "POM|6|A15C|18" & vbCrLf & "POM|7|A15C|23" & vbCrLf & "POM|3|C15D|9" & vbCrLf & "POM|3|C15D|10" & vbCrLf & "POM|4|C15D|12" & vbCrLf & "POM|5|C15D|13" & vbCrLf & "POM|5|C15D|14" & vbCrLf & "POM|3|D15D|9" & vbCrLf & "POM|3|D15D|10" & vbCrLf & "POM|4|D15D|12" & vbCrLf & "POM|5|D15D|13" & vbCrLf & "POM|5|D15D|14" & vbCrLf & "POM|6|G15C|24" & vbCrLf & "POM|7|G15C|28" & vbCrLf & "POM|6|G15D|24" & vbCrLf & "POM|7|G15D|28" & vbCrLf & "POM|6|H15C|24" & vbCrLf & "POM|7|H15C|28" & vbCrLf & "POM|6|H15D|24" & vbCrLf & "POM|7|H15D|28" & vbCrLf & _
    "STR|1|A15C|18" & vbCrLf & "STR|2|A15C|23" & vbCrLf & "STR|3|A15C|27" & vbCrLf & "STR|3|A15C|32" & vbCrLf & "STR|4|A15C|36" & vbCrLf & "STR|5|A15C|40" & vbCrLf & "STR|6|A15C|48" & vbCrLf & "STR|7|A15C|56" & vbCrLf & "STR|8|A15C|64" & vbCrLf & "STR|9|A15C|72" & vbCrLf & "STR|5|D15C|50" & vbCrLf & "STR|6|D15C|55" & vbCrLf & "STR|7|D15C|60" & vbCrLf & "STR|8|D15C|65" & vbCrLf & "STR|5|D15D|50" & vbCrLf & "STR|6|D15D|55" & vbCrLf & "STR|7|D15D|60" & vbCrLf & "STR|8|D15D|65" & vbCrLf & "STR|1|E15C|28" & vbCrLf & "STR|2|E15C|32" & vbCrLf & "STR|3|E15C|35" & vbCrLf & "STR|3|E15C|40" & vbCrLf & "STR|4|E15C|45" & vbCrLf & "STR|5|E15C|50" & vbCrLf & "STR|6|E15C|55" & vbCrLf & _
    "STR|7|E15C|60" & vbCrLf & "STR|8|E15C|65" & vbCrLf & "STR|1|E15D|28" & vbCrLf & "STR|2|E15D|32" & vbCrLf & "STR|3|E15D|35" & vbCrLf & "STR|3|E15D|40" & vbCrLf & "STR|4|E15D|45" & vbCrLf & "STR|5|E15D|50" & vbCrLf & "STR|6|E15D|55" & vbCrLf & "STR|7|E15D|60" & vbCrLf & "STR|8|E15D|65" & vbCrLf & "STR|1|G15C|28" & vbCrLf & "STR|2|G15C|32" & vbCrLf & "STR|1|H15C|28" & vbCrLf & "STR|2|H15C|32" & vbCrLf & "STR|1|G15D|28" & vbCrLf & "STR|2|G15D|32" & vbCrLf & "STR|1|H15D|28" & vbCrLf & "STR|2|H15D|32" & vbCrLf & "STR|1|J60B|L" & vbCrLf & "STR|2|J60B|M" & vbCrLf & "STR|3|J60B|S" & vbCrLf & "STR|1|Z10D|18" & vbCrLf & "STR|2|Z10D|20" & vbCrLf & "STR|3|Z10D|22" & vbCrLf & _
    "VAL|0|A15C|32" & vbCrLf & "VAL|1|A15C|36" & vbCrLf & "VAL|2|A15C|40" & vbCrLf & "VAL|3|A15C|48" & vbCrLf & "VAL|4|A15C|56" & vbCrLf & "VAL|5|A15C|64" & vbCrLf & "VAL|6|A15C|72" & vbCrLf & "VAL|7|A15C|88" & vbCrLf & "VAL|8|A15C|105" & vbCrLf & "VAL|9|A15C|125" & vbCrLf & "VAL|0|E15D|40" & vbCrLf & "VAL|1|E15D|45" & vbCrLf & "VAL|2|E15D|50" & vbCrLf & "VAL|3|E15D|55" & vbCrLf & "VAL|4|E15D|60" & vbCrLf & "VAL|5|E15D|65" & vbCrLf & "VAL|6|E15D|72" & vbCrLf & "VAL|7|E15D|90" & vbCrLf & "VAL|8|E15D|96" & vbCrLf & "VAL|0|E15C|40" & vbCrLf & "VAL|1|E15C|45" & vbCrLf & "VAL|2|E15C|50" & vbCrLf & "VAL|3|E15C|55" & vbCrLf & "VAL|4|E15C|60" & vbCrLf & "VAL|5|E15C|65" & vbCrLf & "VAL|6|E15C|72" & vbCrLf & "VAL|7|E15C|90" & vbCrLf & _
    "VAL|8|E15C|96" & vbCrLf & "VAL|0|D15D|40" & vbCrLf & "VAL|1|D15D|45" & vbCrLf & "VAL|2|D15D|50" & vbCrLf & "VAL|3|D15D|55" & vbCrLf & "VAL|4|D15D|60" & vbCrLf & "VAL|5|D15D|65" & vbCrLf & "VAL|6|D15D|72" & vbCrLf & "VAL|7|D15D|90" & vbCrLf & "VAL|8|D15D|96" & vbCrLf & "VAL|0|C15D|40" & vbCrLf & "VAL|1|C15D|45" & vbCrLf & "VAL|2|C15D|50" & vbCrLf & "VAL|3|C15D|55" & vbCrLf & "VAL|4|C15D|60" & vbCrLf & "VAL|5|C15D|65" & vbCrLf & "VAL|6|C15D|72" & vbCrLf & "VAL|7|C15D|90" & vbCrLf & "VAL|8|C15D|96" & vbCrLf & "VAL|0|Z10D|20" & vbCrLf & "VAL|0|Z10D|22" & vbCrLf & "VAL|1|Z10D|23" & vbCrLf & "VAL|2|Z10D|25" & vbCrLf & "VAL|3|Z10D|28" & vbCrLf & "VAL|4|Z10D|30" & vbCrLf & "VAL|5|Z10D|33" & vbCrLf & _
    "VAL|0|D10D|20" & vbCrLf & "VAL|1|D10D|24" & vbCrLf & "VAL|2|D10D|28" & vbCrLf & "VAL|3|D10D|32" & vbCrLf & "VAL|4|D10D|36" & vbCrLf & "VAL|5|D10D|40" & vbCrLf & "VAL|0|E10D|20" & vbCrLf & "VAL|1|E10D|24" & vbCrLf & "VAL|2|E10D|28" & vbCrLf & "VAL|3|E10D|32" & vbCrLf & "VAL|4|E10D|36" & vbCrLf & "VAL|5|E10D|40"


Private mSizeTable As Object   ' "VARIETY|PACK" -> Collection of Count strings
Private mRefTable As Object    ' "VARIETY|PACK|SIZEREF" -> Collection of Count strings

' ---------------------------------------------------------------
' The Variety column (D:D) is no longer read at all. It can now
' contain ambiguous placeholder text like "Any seedless" or "Any
' variety" and it's simply ignored. Instead, the family (VAL, STR,
' POM, or LEM) is detected purely from the numbers already present
' in the header cell itself, scored against each family's known
' counts across ALL of its packs (not just the leading token).
'
' WHY SCORING ACROSS THE WHOLE HEADER, NOT JUST THE LEADING TOKEN:
' VAL and STR's A15C count ranges overlap heavily (32, 36, 40, 48,
' 56, 64, 72 are valid A15C counts for BOTH families), so the leading
' token alone often can't tell them apart. But it's very unlikely
' two families would coincidentally explain the ENTIRE set of
' numbers in one header cell the same way, so each candidate family
' is scored by how many of the header's bracket tokens it can
' account for via its own packs, and the best-explained family wins.
'
' TIE-BREAK: if two families score identically (rare), preference
' order is VAL, then STR, then POM, then LEM. Adjust FAMILY_ORDER
' below if that priority ever needs to change.
' ---------------------------------------------------------------
Private Const FAMILY_ORDER As String = "VAL,STR,POM,LEM"
' ---------------------------------------------------------------
' Per-row family cache, so DetectFamilyForRow only has to scan the
' row's header cells once, no matter how many columns call
' GetSizeCriteria against that same row afterwards.
' Key: shName & "|" & startline  ->  detected family (String)
' ---------------------------------------------------------------
Private mRowFamilyCache As Object

' ---------------------------------------------------------------
' Every distinct pack code that exists anywhere in SIZE_DATA for a
' given family.
' ---------------------------------------------------------------
Private Function DistinctPacksForVariety(ByVal variety As String) As Collection
    Dim result As New Collection
    Dim k As Variant, parts() As String
    For Each k In mSizeTable.Keys
        parts = Split(CStr(k), "|")
        If parts(0) = variety Then result.Add parts(1)
    Next k
    Set DistinctPacksForVariety = result
End Function

' ---------------------------------------------------------------
' How many of the bracket tokens can this family explain, using
' every pack EXCEPT the ones that already own the leading token?
' Z10D uses the broad "anywhere in its count list" rule (matching
' its real resolution behavior); every other pack must match at the
' exact SIZE_REF given.
' ---------------------------------------------------------------
Private Function CountExplainedTokens(ByVal variety As String, ByVal ref As Long, _
                                       ByVal bracketTokens As Collection, ByVal excludePacks As Collection) As Long
    Dim explainedCount As Long, t As Variant, p As Variant, explained As Boolean
    Dim cands As Collection

    explainedCount = 0
    For Each t In bracketTokens
        explained = False
        For Each p In DistinctPacksForVariety(variety)
            If Not TokenInCollection(CStr(p), excludePacks) Then
                If CStr(p) = "Z10D" Then
                    If mSizeTable.Exists(variety & "|Z10D") Then
                        If TokenInCollection(CStr(t), mSizeTable(variety & "|Z10D")) Then explained = True
                    End If
                Else
                    Set cands = CountsAtRef(variety, CStr(p), ref)
                    If TokenInCollection(CStr(t), cands) Then explained = True
                End If
            End If
            If explained Then Exit For
        Next p
        If explained Then explainedCount = explainedCount + 1
    Next t

    CountExplainedTokens = explainedCount
End Function

' ---------------------------------------------------------------
' Tries every known family against this header's tokens and returns
' whichever one best explains the whole header - not just the
' leading token. Returns "" if no family recognizes the leading
' token at all (e.g. an X-code header).
' ---------------------------------------------------------------
Private Function DetectFamily(ByVal leadingToken As String, ByVal bracketTokens As Collection, _
                               ByRef outRef As Long, ByRef outLeadingFamily As Collection) As String
    Dim families() As String
    families = Split(FAMILY_ORDER, ",")

    Dim bestFamily As String, bestScore As Long, bestRef As Long
    Dim bestLeadingFamily As Collection
    bestScore = -1
    bestFamily = ""

    Dim i As Long, fam As String, lf As Collection, r As Long, score As Long
    For i = LBound(families) To UBound(families)
        fam = Trim(families(i))
        Set lf = DetermineLeadingFamily(fam, leadingToken)
        If lf.Count > 0 Then
            r = FindSizeRefForPackCount(fam, CStr(lf(1)), leadingToken)
            score = CountExplainedTokens(fam, r, bracketTokens, lf)
            If score > bestScore Then
                bestScore = score
                bestFamily = fam
                bestRef = r
                Set bestLeadingFamily = lf
            End If
        End If
    Next i

    outRef = bestRef
    Set outLeadingFamily = bestLeadingFamily
    DetectFamily = bestFamily
End Function

' ---------------------------------------------------------------
' Call this manually (e.g. before a full recalculation) if header
' text on the sheet has changed and you want to force the row-level
' family detection to run again instead of using cached results.
' ---------------------------------------------------------------
Public Sub ClearRowFamilyCache()
    Set mRowFamilyCache = Nothing
End Sub

' ---------------------------------------------------------------
' Detects the family for an ENTIRE header row at once, by scoring
' every header cell from colFrom to colTo together, rather than one
' column in isolation. A sparse column like "36" alone is often
' ambiguous (VAL and STR both have 36 as an A15C count), but if
' other columns in the SAME row clearly favor one family (e.g. via
' Z10D or E-group brackets that only make sense for VAL), that
' evidence should carry over to the sparse column too - the whole
' row is the same variety, after all.
'
' Blank cells contribute nothing and are simply skipped.
' ---------------------------------------------------------------
Private Function DetectFamilyForRow(ByVal shName As String, ByVal startline As Long, _
                                     ByVal colFrom As Long, ByVal colTo As Long) As String
    If mRowFamilyCache Is Nothing Then Set mRowFamilyCache = CreateObject("Scripting.Dictionary")

    Dim cacheKey As String
    cacheKey = shName & "|" & startline
    If mRowFamilyCache.Exists(cacheKey) Then
        DetectFamilyForRow = mRowFamilyCache(cacheKey)
        Exit Function
    End If

    Dim families() As String
    families = Split(FAMILY_ORDER, ",")

    Dim totalScore() As Long
    ReDim totalScore(LBound(families) To UBound(families))

    Dim col As Long, headerText As String, tokens As Collection
    Dim leadTok As String, bracketToks As New Collection, k As Long

    For col = colFrom To colTo
        headerText = Trim(ThisWorkbook.Worksheets(shName).Cells(startline, col).Value)
        If Len(headerText) > 0 Then
            Set tokens = ParseHeaderTokens(headerText)
            leadTok = CStr(tokens(1))
            Set bracketToks = New Collection
            For k = 2 To tokens.Count
                bracketToks.Add CStr(tokens(k))
            Next k

            Dim i As Long, fam As String, lf As Collection, r As Long
            For i = LBound(families) To UBound(families)
                fam = Trim(families(i))
                Set lf = DetermineLeadingFamily(fam, leadTok)
                If lf.Count > 0 Then
                    r = FindSizeRefForPackCount(fam, CStr(lf(1)), leadTok)
                    ' +1 for the leading token itself validating this
                    ' family, plus however many bracket tokens it
                    ' explains - summed across every column in the row.
                    totalScore(i) = totalScore(i) + 1 + CountExplainedTokens(fam, r, bracketToks, lf)
                End If
            Next i
        End If
    Next col

    Dim bestIdx As Long, bestScore As Long
    bestScore = -1
    For i = LBound(families) To UBound(families)
        If totalScore(i) > bestScore Then
            bestScore = totalScore(i)
            bestIdx = i
        End If
    Next i

    If bestScore <= 0 Then
        DetectFamilyForRow = ""
    Else
        DetectFamilyForRow = Trim(families(bestIdx))
    End If

    mRowFamilyCache.Add cacheKey, DetectFamilyForRow
End Function

' ---------------------------------------------------------------
' Build the lookup table once from SIZE_DATA.
' ---------------------------------------------------------------
Private Sub EnsureSizeTable()
    If Not mSizeTable Is Nothing Then Exit Sub
    Set mSizeTable = CreateObject("Scripting.Dictionary")
    Set mRefTable = CreateObject("Scripting.Dictionary")

    Dim lines() As String, f() As String
    Dim i As Long, key As String, refKey As String
    lines = Split(SIZE_DATA, vbCrLf)

    For i = LBound(lines) To UBound(lines)
        If Len(Trim(lines(i))) > 0 Then
            f = Split(lines(i), "|")
            key = f(0) & "|" & f(2)                      ' VARIETY|PACK
            If Not mSizeTable.Exists(key) Then
                mSizeTable.Add key, New Collection
            End If
            mSizeTable(key).Add f(3)                       ' Count

            refKey = f(0) & "|" & f(2) & "|" & f(1)        ' VARIETY|PACK|SIZEREF
            If Not mRefTable.Exists(refKey) Then
                mRefTable.Add refKey, New Collection
            End If
            mRefTable(refKey).Add f(3)                     ' Count
        End If
    Next i
End Sub

' ---------------------------------------------------------------
' All counts for a given Variety+Pack at one exact SIZE_REF. Usually
' a single value; occasionally more than one (e.g. LEM/A15C SizeRef3
' has both 88 and 100) - a genuine pre-existing data ambiguity, not
' something this function tries to resolve.
' ---------------------------------------------------------------
Private Function CountsAtRef(ByVal variety As String, ByVal pack As String, ByVal ref As Long) As Collection
    Dim key As String
    key = variety & "|" & pack & "|" & CStr(ref)
    If mRefTable.Exists(key) Then
        Set CountsAtRef = mRefTable(key)
    Else
        Set CountsAtRef = New Collection
    End If
End Function

' ---------------------------------------------------------------
' Finds the SIZE_REF at which Variety+Pack has the given count value.
' Returns -1 if not found. If the count appears at more than one ref
' (shouldn't normally happen), the first one encountered is used.
' ---------------------------------------------------------------
Private Function FindSizeRefForPackCount(ByVal variety As String, ByVal pack As String, ByVal countVal As String) As Long
    Dim lines() As String, f() As String, i As Long
    lines = Split(SIZE_DATA, vbCrLf)

    FindSizeRefForPackCount = -1
    For i = LBound(lines) To UBound(lines)
        If Len(Trim(lines(i))) > 0 Then
            f = Split(lines(i), "|")
            If f(0) = variety And f(2) = pack And f(3) = countVal Then
                FindSizeRefForPackCount = CLng(f(1))
                Exit Function
            End If
        End If
    Next i
End Function

' ---------------------------------------------------------------
' Splits a header cell like "36 (20,23,45)" or "36(45)" or "72"
' into a Collection of trimmed tokens: ("36","20","23","45").
' Handles any combination of spaces before "(" and after ",".
' ---------------------------------------------------------------
Private Function ParseHeaderTokens(ByVal headerText As String) As Collection
    Dim result As New Collection
    Dim cleaned As String, leading As String, inside As String
    Dim posParen As Long, parts() As String, i As Long

    cleaned = Trim(headerText)
    posParen = InStr(cleaned, "(")

    If posParen = 0 Then
        result.Add cleaned
    Else
        leading = Trim(Left(cleaned, posParen - 1))
        inside = Mid(cleaned, posParen + 1)
        inside = Replace(inside, ")", "")
        result.Add leading
        parts = Split(inside, ",")
        For i = LBound(parts) To UBound(parts)
            If Len(Trim(parts(i))) > 0 Then result.Add Trim(parts(i))
        Next i
    End If

    Set ParseHeaderTokens = result
End Function

Private Function TokenInCollection(ByVal tok As String, ByVal col As Collection) As Boolean
    Dim v As Variant
    For Each v In col
        If CStr(v) = tok Then
            TokenInCollection = True
            Exit Function
        End If
    Next v
    TokenInCollection = False
End Function

Private Function IsNumericToken(ByVal s As String) As Boolean
    IsNumericToken = IsNumeric(s)
End Function

Private Function FormatToken(ByVal tok As String) As String
    If IsNumericToken(tok) Then
        FormatToken = tok
    Else
        FormatToken = Chr(34) & tok & Chr(34)
    End If
End Function

' ---------------------------------------------------------------
' The repeated COUNTIFS criteria block, identical every time except
' the trailing $M:$M count and the $P:$P criteria (which depends on
' the ansName!V1 "OFF" toggle) - mirrors the original "20,22" case.
' ---------------------------------------------------------------
' ---------------------------------------------------------------
' The criteria that NEVER vary between combo terms: C, P(D or
' literal "VAL"), E, F, I, J, K, R, L. Brand (G), Pack (H), and the
' size count (M) are deliberately excluded here - they now vary per
' combo term because any of the three can be compound (e.g. brand
' "MALA/MAHEL", pack "E10D/D10D"), so BuildFullFormula assembles
' those separately for each term in the cross-product.
' ---------------------------------------------------------------
Private Function FixedCriteria(ByVal ansName As String, ByVal vals As Long, ByVal useLiteralVAL As Boolean) As String
    Dim q As String
    q = Chr(34)

    Dim pCriteria As String
    If useLiteralVAL Then
        pCriteria = ansName & "!$P:$P," & q & "VAL" & q & ","
    Else
        pCriteria = ansName & "!$P:$P,$D" & vals & ","
    End If

    FixedCriteria = _
        ansName & "!$C:$C,$C" & vals & "," & _
        pCriteria & _
        ansName & "!$E:$E,$E" & vals & "," & _
        ansName & "!$F:$F,$F" & vals & "," & _
        ansName & "!$I:$I,IF(ISBLANK($I" & vals & ")," & q & q & ",$I" & vals & ")," & _
        ansName & "!$J:$J,$J" & vals & "," & _
        ansName & "!$K:$K,$K" & vals & "," & _
        ansName & "!$R:$R," & q & q & "," & _
        ansName & "!$L:$L,IF(ISBLANK($P" & vals & ")," & q & q & ",$P" & vals & "),"
End Function

' ---------------------------------------------------------------
' Resolves the size COUNT value(s) for ONE individual pack code
' (never a compound string - the caller splits "E10D/D10D" etc.
' before calling this per pack). Returns raw values, not formula
' text - BuildFullFormula assembles the actual COUNTIFS terms.
'
' Same resolution rules as before: the pack owning the leading token
' returns it directly; Z10D uses the broad "anywhere in its count
' list" rule (can return more than one value, meaning "sum both");
' every other pack uses the exact SIZE_REF lookup, falling back to
' the leading token if it has no distinct entry there.
' ---------------------------------------------------------------
Private Function ResolveCountsForPack(ByVal variety As String, ByVal pack As String, ByVal ref As Long, _
                                       ByVal leadingToken As String, ByVal bracketTokens As Collection, _
                                       ByVal leadingFamily As Collection) As Collection
    Dim result As New Collection
    Dim v As Variant

    If TokenInCollection(pack, leadingFamily) Then
        result.Add leadingToken
        Set ResolveCountsForPack = result
        Exit Function
    End If

    If pack = "Z10D" Then
        If mSizeTable.Exists(variety & "|Z10D") Then
            For Each v In mSizeTable(variety & "|Z10D")
                If TokenInCollection(CStr(v), bracketTokens) Then result.Add CStr(v)
            Next v
        End If
        If result.Count = 0 Then result.Add leadingToken
        Set ResolveCountsForPack = result
        Exit Function
    End If

    If ref >= 0 Then
        Dim candidates As Collection
        Set candidates = CountsAtRef(variety, pack, ref)
        If candidates.Count = 1 Then
            result.Add CStr(candidates(1))
            Set ResolveCountsForPack = result
            Exit Function
        ElseIf candidates.Count > 1 Then
            Dim disambiguated As New Collection
            For Each v In candidates
                If TokenInCollection(CStr(v), bracketTokens) Then disambiguated.Add CStr(v)
            Next v
            If disambiguated.Count >= 1 Then
                result.Add CStr(disambiguated(1))
                Set ResolveCountsForPack = result
                Exit Function
            End If
        End If
    End If

    result.Add leadingToken
    Set ResolveCountsForPack = result
End Function

' ---------------------------------------------------------------
' Canonical key representing a count SET (order-independent), used
' to detect which packs are numerically indistinguishable for a
' given variety (e.g. E15D/E15C/D15D/C15D share identical counts).
' ---------------------------------------------------------------
Private Function CountSetKey(ByVal counts As Collection) As String
    Dim arr() As String, i As Long, a As Long, b As Long, tmp As String
    ReDim arr(1 To counts.Count)
    For i = 1 To counts.Count
        arr(i) = CStr(counts(i))
    Next i
    For a = 1 To UBound(arr) - 1
        For b = a + 1 To UBound(arr)
            If arr(b) < arr(a) Then
                tmp = arr(a): arr(a) = arr(b): arr(b) = tmp
            End If
        Next b
    Next a
    CountSetKey = Join(arr, "|")
End Function

' ---------------------------------------------------------------
' All packs (for a variety) whose count set exactly matches `pack`'s.
' ---------------------------------------------------------------
Private Function PacksInSameFamily(ByVal variety As String, ByVal pack As String) As Collection
    Dim result As New Collection
    Dim targetKey As String, k As Variant, parts() As String

    If Not mSizeTable.Exists(variety & "|" & pack) Then
        result.Add pack
        Set PacksInSameFamily = result
        Exit Function
    End If

    targetKey = CountSetKey(mSizeTable(variety & "|" & pack))
    For Each k In mSizeTable.Keys
        parts = Split(CStr(k), "|")
        If parts(0) = variety Then
            If CountSetKey(mSizeTable(k)) = targetKey Then result.Add parts(1)
        End If
    Next k

    Set PacksInSameFamily = result
End Function

' ---------------------------------------------------------------
' Works out which pack-family "owns" the leading (outside-bracket)
' token for this header, for this variety. A15C is preferred
' whenever the leading token is a valid A15C count (the normal,
' by-far-most-common case). If it ISN'T a valid A15C count, but IS
' valid for some other family (e.g. E15D), that family is used
' instead - this is what lets a rare "E15D leads" header resolve
' correctly without hardcoding it.
'
' KNOWN LIMITATION: if the leading token happens to be a value valid
' for BOTH A15C and another family (this does happen - e.g. "40" and
' "72" are valid A15C counts AND valid E-group counts for VAL), this
' defaults to A15C. There's no way to disambiguate from the number
' alone in that case; it would need an extra signal on the sheet.
' ---------------------------------------------------------------
Private Function DetermineLeadingFamily(ByVal variety As String, ByVal leadingToken As String) As Collection
    Dim result As New Collection

    If TokenInCollection(leadingToken, mSizeTable(variety & "|A15C")) Then
        Set DetermineLeadingFamily = PacksInSameFamily(variety, "A15C")
        Exit Function
    End If

    Dim k As Variant, parts() As String
    Dim seenKeys As Object
    Set seenKeys = CreateObject("Scripting.Dictionary")

    For Each k In mSizeTable.Keys
        parts = Split(CStr(k), "|")
        If parts(0) = variety Then
            If TokenInCollection(leadingToken, mSizeTable(k)) Then
                Dim famKey As String
                famKey = CountSetKey(mSizeTable(k))
                If Not seenKeys.Exists(famKey) Then
                    seenKeys.Add famKey, True
                    Set result = PacksInSameFamily(variety, parts(1))
                    Set DetermineLeadingFamily = result
                    Exit Function
                End If
            End If
        End If
    Next k

    ' Nothing in the table matches the leading token at all - leave
    ' the caller to fall back to it directly regardless of pack.
    Set DetermineLeadingFamily = New Collection
End Function

' ---------------------------------------------------------------
' MAIN ENTRY POINT
' Builds the ENTIRE formula (from "=IF(" to the final close),
' including the fixed criteria, the size lookup, AND the brand/pack
' handling - not just the size portion. Call it directly in place of
' building sForm by hand:
'
'     ThisWorkbook.Worksheets(shName).Cells(i, j).Formula = _
'         BuildFullFormula(shName, ansName, startline, j, vals, sPack, , totCol)
'
' WHY THIS NOW OWNS THE WHOLE FORMULA, NOT JUST THE SIZE PART:
' Brand (Cells(vals,7)) and Pack (Cells(vals,8)) can now ALSO be
' compound, e.g. "MALA/MAHEL" or "E10D/D10D" - the packhouse is
' deliberately mixing brands/packs on one line to work through
' stock. A compound value means "count rows matching ANY of these",
' which is the same idea as summing multiple size counts, but now
' three dimensions (brand, pack, size) can all vary at once, and
' every valid COMBINATION needs its own COUNTIFS term, summed
' together. Pack and size are resolved TOGETHER per pack (since each
' individual pack can have its own count), then brand multiplies
' across whatever pack/count pairs result.
'
' BEHAVIOUR CHANGE FROM THE OLD SPLIT ASSEMBLY - PLEASE CONFIRM:
' The dispatch-flag criterion ($Q:$Q,1) used to be added only to the
' FIRST summed term (inside the old size-only helper) and separately
' to whatever ended up being the LAST term (via the calling Sub's own
' closing code) - meaning any MIDDLE terms silently missed it
' whenever there were 3 or more. That was invisible with at most two
' terms; it won't stay invisible with brand x pack x size
' combinations. This version applies $Q:$Q,1 to EVERY term when
' sPack <> "PACKHOUSE", uniformly. If the old first/last-only
' behaviour was actually intentional for some reason, tell me and
' I'll revert this specific piece.
' ---------------------------------------------------------------
Public Function BuildFullFormula(ByVal shName As String, ByVal ansName As String, ByVal startline As Long, _
                                  ByVal j As Long, ByVal vals As Long, ByVal sPack As String, _
                                  Optional ByVal colFrom As Long = 17, Optional ByVal colTo As Long = 0) As String
    EnsureSizeTable

    Dim headerText As String, tokens As Collection, bracketTokens As New Collection
    Dim leadingToken As String, i As Long

    headerText = ThisWorkbook.Worksheets(shName).Cells(startline, j).Value
    Set tokens = ParseHeaderTokens(headerText)
    leadingToken = CStr(tokens(1))
    For i = 2 To tokens.Count
        bracketTokens.Add CStr(tokens(i))
    Next i

    ' --- Determine family + SIZE_REF (unchanged from before) ---
    Dim variety As String, ref As Long, leadingFamily As Collection
    If colFrom > 0 And colTo >= colFrom Then
        variety = DetectFamilyForRow(shName, startline, colFrom, colTo)
    Else
        variety = DetectFamily(leadingToken, bracketTokens, ref, leadingFamily)
    End If

    If variety <> "" Then
        Set leadingFamily = DetermineLeadingFamily(variety, leadingToken)
        If leadingFamily.Count > 0 Then
            ref = FindSizeRefForPackCount(variety, CStr(leadingFamily(1)), leadingToken)
        Else
            ref = -1
        End If
    Else
        Set leadingFamily = New Collection
        ref = -1
    End If

    ' --- Large/Medium/Small marker: forces the size to "L"/"M"/"S"
    '     for every combo term, but brand/pack still expand normally ---
    Dim forcedSize As String
    forcedSize = ""
    Select Case ThisWorkbook.Worksheets(shName).Cells(vals, j).Value
        Case "Large":  forcedSize = "L"
        Case "Medium": forcedSize = "M"
        Case "Small":  forcedSize = "S"
    End Select

    ' --- Split Brand and Pack into individual values ---
    Dim brandList() As String, packList() As String
    brandList = Split(Trim(ThisWorkbook.Worksheets(shName).Cells(vals, 7).Value), "/")
    packList = Split(Trim(ThisWorkbook.Worksheets(shName).Cells(vals, 8).Value), "/")

    ' --- Resolve size count(s) for EACH individual pack first, since
    '     pack and size are linked (each pack can have its own count) ---
    Dim pcPacks As New Collection, pcCounts As New Collection
    Dim p As Variant, thisPack As String, cnts As Collection, c As Variant

    For Each p In packList
        thisPack = Trim(CStr(p))
        If Len(forcedSize) > 0 Then
            pcPacks.Add thisPack
            pcCounts.Add forcedSize
        Else
            Set cnts = ResolveCountsForPack(variety, thisPack, ref, leadingToken, bracketTokens, leadingFamily)
            For Each c In cnts
                pcPacks.Add thisPack
                pcCounts.Add CStr(c)
            Next c
        End If
    Next p

    ' --- Cross the (pack,count) pairs with every brand value ---
    Dim termBrand As New Collection, termPack As New Collection, termCount As New Collection
    Dim bIdx As Long, pcIdx As Long, brandVal As String

    For bIdx = LBound(brandList) To UBound(brandList)
        brandVal = Trim(brandList(bIdx))
        For pcIdx = 1 To pcPacks.Count
            termBrand.Add brandVal
            termPack.Add pcPacks(pcIdx)
            termCount.Add pcCounts(pcIdx)
        Next pcIdx
    Next bIdx

    ' --- Assemble the full formula ---
    Dim useLiteralVAL As Boolean
    useLiteralVAL = (ThisWorkbook.Worksheets(ansName).Range("V1").Value <> "OFF")

    Dim fixedCrit As String
    fixedCrit = FixedCriteria(ansName, vals, useLiteralVAL)

    Dim q As String
    q = Chr(34)

    Dim formulaBody As String
    Dim t As Long

    For t = 1 To termBrand.Count
        formulaBody = formulaBody & fixedCrit & _
            ansName & "!$G:$G," & q & CStr(termBrand(t)) & q & "," & _
            ansName & "!$H:$H," & q & CStr(termPack(t)) & q & "," & _
            ansName & "!$M:$M," & FormatToken(CStr(termCount(t)))

        If sPack <> "PACKHOUSE" Then
            formulaBody = formulaBody & ", " & ansName & "!$Q:$Q,1"
        End If

        If t < termBrand.Count Then
            formulaBody = formulaBody & ")+COUNTIFS("
        End If
    Next t

    Dim colLetter As String
    colLetter = charCheck(j + 64)

    Dim result As String
    result = "=IF(NOT(ISBLANK(" & colLetter & vals & ")),COUNTIFS(" & formulaBody & ")"

    If sPack = "PACKHOUSE" Then
        result = result & ",0)-" & colLetter & (vals + 5)
    Else
        result = result & ",0)"
    End If

    BuildFullFormula = result
End Function

