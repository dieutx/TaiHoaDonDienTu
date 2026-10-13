Attribute VB_Name = "modGdtResume"
Option Explicit

' Resume re-queries the selected period; cursors and credentials are never saved.
' Only rows owned by the same account/source can be reused or replaced.
Private Const STATE_SHEET As String = "TrangThaiTaiHD"
Private mAccount As String, mDirection As String
Private mRows As Object, mSeen As Object, mSignatures As Object, mStages As Object
Private mDetailCounts As Object, mReplaced As Object
Private mLegacyRows As Object, mLegacyDetails As Object, mLegacyReplaced As Object
Private mOriginalDetailLast As Long
Public GdtResumeEnabled As Boolean
Public GdtDuplicatesSkipped As Long, GdtStagesSkipped As Long

Public Function GdtResumeSessionActive() As Boolean
    GdtResumeSessionActive = Not mSeen Is Nothing
End Function

Private Function KeyPart(ByVal value As String) As String
    KeyPart = CStr(Len(value)) & ":" & value
End Function

Public Function GdtInvoiceIdentity(ByVal seller As String, ByVal template As String, _
    ByVal series As String, ByVal number As String) As String
    GdtInvoiceIdentity = KeyPart(seller) & KeyPart(template) & KeyPart(series) & KeyPart(number)
End Function

Private Function ScopeKey(ByVal identity As String, ByVal source As String) As String
    ScopeKey = KeyPart(mAccount) & KeyPart(mDirection) & KeyPart(source) & identity
End Function

Public Sub BeginGdtResume(ByVal account As String, ByVal invoiceType As Long, ByVal appendExisting As Boolean)
    Dim ws As Worksheet, state As Worksheet, data As Variant, r As Long, lastRow As Long, identity As String
    Set state = GetStateSheet(False)
    ValidateOwnerColumns ThisWorkbook.Sheets(IIf(invoiceType = 1, "TongHopHD_Mua", "TongHopHD_Ban")), 65
    ValidateOwnerColumns ThisWorkbook.Sheets(IIf(invoiceType = 1, "ChiTietHD_Mua", "ChiTietHD_Ban")), 35
    mAccount = account: mDirection = CStr(invoiceType)
    GdtResumeEnabled = appendExisting
    GdtDuplicatesSkipped = 0: GdtStagesSkipped = 0
    Set mRows = CreateObject("Scripting.Dictionary")
    Set mSeen = CreateObject("Scripting.Dictionary")
    Set mSignatures = CreateObject("Scripting.Dictionary")
    Set mStages = CreateObject("Scripting.Dictionary")
    Set mDetailCounts = CreateObject("Scripting.Dictionary")
    Set mReplaced = CreateObject("Scripting.Dictionary")
    Set mLegacyRows = CreateObject("Scripting.Dictionary")
    Set mLegacyDetails = CreateObject("Scripting.Dictionary")
    Set mLegacyReplaced = CreateObject("Scripting.Dictionary")
    Set ws = ThisWorkbook.Sheets(IIf(invoiceType = 1, "TongHopHD_Mua", "TongHopHD_Ban"))
    lastRow = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    If appendExisting And lastRow >= 3 Then
        data = ws.Cells(3, 1).Resize(lastRow - 2, 66).Value2
        For r = 1 To UBound(data, 1)
            If CStr(data(r, 65)) = account And Len(CStr(data(r, 66))) > 0 Then
                identity = GdtInvoiceIdentity(CStr(data(r, 10)), CStr(data(r, 3)), CStr(data(r, 4)), CStr(data(r, 5)))
                If Not mRows.Exists(ScopeKey(identity, CStr(data(r, 66)))) Then _
                    mRows.Add ScopeKey(identity, CStr(data(r, 66))), r + 2
            ElseIf Len(CStr(data(r, 65))) = 0 And Len(account) > 0 Then
                If CStr(data(r, IIf(invoiceType = 1, 16, 10))) = account Then
                    identity = GdtInvoiceIdentity(CStr(data(r, 10)), CStr(data(r, 3)), CStr(data(r, 4)), CStr(data(r, 5)))
                    If Not mLegacyRows.Exists(identity) Then mLegacyRows(identity) = r + 2
                End If
            End If
        Next r
    End If
    Set ws = ThisWorkbook.Sheets(IIf(invoiceType = 1, "ChiTietHD_Mua", "ChiTietHD_Ban"))
    mOriginalDetailLast = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    If mOriginalDetailLast >= 3 Then
        data = ws.Cells(3, 1).Resize(mOriginalDetailLast - 2, 36).Value2
        For r = 1 To UBound(data, 1)
            If CStr(data(r, 35)) = account And Len(CStr(data(r, 36))) > 0 Then
                identity = ScopeKey(GdtInvoiceIdentity(CStr(data(r, 8)), CStr(data(r, 1)), CStr(data(r, 2)), CStr(data(r, 3))), CStr(data(r, 36)))
                If Not mDetailCounts.Exists(identity) Then mDetailCounts(identity) = 0
                mDetailCounts(identity) = CLng(mDetailCounts(identity)) + 1
            ElseIf Len(CStr(data(r, 35))) = 0 And Len(account) > 0 Then
                If CStr(data(r, IIf(invoiceType = 1, 14, 8))) = account Then
                    identity = GdtInvoiceIdentity(CStr(data(r, 8)), CStr(data(r, 1)), CStr(data(r, 2)), CStr(data(r, 3)))
                    mLegacyDetails(identity) = True
                End If
            End If
        Next r
    End If
    If Not state Is Nothing Then
        lastRow = state.Cells(state.Rows.Count, 1).End(xlUp).Row
        If lastRow >= 2 Then
            data = state.Cells(2, 1).Resize(lastRow - 1, 6).Value2
            For r = 1 To UBound(data, 1)
                If Len(CStr(data(r, 1))) > 0 And IsNumeric(data(r, 4)) Then
                    If CDbl(data(r, 4)) >= 0 And CDbl(data(r, 4)) <= ws.Rows.Count Then _
                        mStages(CStr(data(r, 1))) = Array(CStr(data(r, 2)), CLng(data(r, 4)), r + 1, CStr(data(r, 6)))
                End If
            Next r
        End If
    End If
End Sub

Private Sub ValidateOwnerColumns(ByVal ws As Worksheet, ByVal firstColumn As Long)
    Dim i As Long, expected As String, lastRow As Long
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    For i = 0 To 1
        expected = IIf(i = 0, "Account scope", "API source")
        If CStr(ws.Cells(2, firstColumn + i).Value2) <> expected Then
            If Application.CountA(ws.Cells(1, firstColumn + i).Resize(lastRow, 1)) > 0 Then _
                Err.Raise 5, , "Resume metadata column is already in use: " & ws.Name & "!" & ws.Cells(2, firstColumn + i).Address(False, False)
        End If
    Next i
End Sub

Public Sub EndGdtResume()
    GdtResumeEnabled = False
    Set mRows = Nothing: Set mSeen = Nothing: Set mSignatures = Nothing: Set mStages = Nothing
    Set mDetailCounts = Nothing: Set mReplaced = Nothing
    Set mLegacyRows = Nothing: Set mLegacyDetails = Nothing: Set mLegacyReplaced = Nothing
End Sub

Public Sub PrepareGdtResumePage(ByVal page As Object, ByVal source As String, ByRef targets As Object)
    Dim accepted As New Collection, invoice As Object, identity As String, key As String, pageSeen As Object
    Set targets = CreateObject("Scripting.Dictionary")
    If mSeen Is Nothing Then Exit Sub
    Set pageSeen = CreateObject("Scripting.Dictionary")
    For Each invoice In page("datas")
        identity = GdtInvoiceIdentity(SafeJsonText(invoice, "nbmst"), SafeJsonText(invoice, "khmshdon"), _
            SafeJsonText(invoice, "khhdon"), SafeJsonText(invoice, "shdon"))
        key = ScopeKey(identity, source)
        If mSeen.Exists(key) Or pageSeen.Exists(key) Then
            GdtDuplicatesSkipped = GdtDuplicatesSkipped + 1
        Else
            accepted.Add invoice
            pageSeen(key) = True
            If GdtResumeEnabled Then
                If mRows.Exists(key) Then
                    targets(identity) = CLng(mRows(key))
                ElseIf mLegacyRows.Exists(identity) Then
                    targets(identity) = CLng(mLegacyRows(identity))
                End If
            End If
        End If
    Next invoice
    Set page("datas") = accepted
End Sub

Public Sub CommitGdtResumePage(ByVal page As Object, ByVal source As String, ByVal targets As Object, _
    ByVal appendStart As Long, ByVal invoiceType As Long)
    Dim invoice As Object, identity As String, key As String, targetRow As Long, ws As Worksheet
    Set ws = ThisWorkbook.Sheets(IIf(invoiceType = 1, "TongHopHD_Mua", "TongHopHD_Ban"))
    ws.Cells(2, 65).Value2 = "Account scope": ws.Cells(2, 66).Value2 = "API source"
    ws.Columns(65).Hidden = True: ws.Columns(66).Hidden = True
    For Each invoice In page("datas")
        identity = GdtInvoiceIdentity(SafeJsonText(invoice, "nbmst"), SafeJsonText(invoice, "khmshdon"), _
            SafeJsonText(invoice, "khhdon"), SafeJsonText(invoice, "shdon"))
        If targets.Exists(identity) Then
            targetRow = CLng(targets(identity))
        Else
            targetRow = appendStart: appendStart = appendStart + 1
        End If
        targets(identity) = targetRow
        If Not mSeen Is Nothing Then
            key = ScopeKey(identity, source)
            mRows(key) = targetRow: mSeen(key) = True
            mSignatures(key) = JsonConverter.ConvertToJson(invoice)
            If mLegacyRows.Exists(identity) Then mLegacyRows.Remove identity
        End If
        GdtWriteTextCell ws.Cells(targetRow, 65), mAccount
        GdtWriteTextCell ws.Cells(targetRow, 66), source
    Next invoice
End Sub

Private Function GetStateSheet(ByVal create As Boolean) As Worksheet
    Dim ws As Worksheet, headers As Variant, i As Long
    headers = Array("Identity (v1)", "Signature", "Stage", "Rows", "Updated", "File proof")
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = STATE_SHEET Then
            For i = 0 To UBound(headers)
                If CStr(ws.Cells(1, i + 1).Value2) <> CStr(headers(i)) Then _
                    Err.Raise 5, , "Unrecognized resume state sheet: " & STATE_SHEET
            Next i
            Set GetStateSheet = ws
            Exit Function
        End If
    Next ws
    If Not create Then Exit Function
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = STATE_SHEET
    ws.Cells(1, 1).Value2 = "Identity (v1)": ws.Cells(1, 2).Value2 = "Signature"
    ws.Cells(1, 3).Value2 = "Stage": ws.Cells(1, 4).Value2 = "Rows": ws.Cells(1, 5).Value2 = "Updated"
    ws.Cells(1, 6).Value2 = "File proof"
    ws.Visible = xlSheetVeryHidden
    Set GetStateSheet = ws
End Function

Private Function BufferKey(ByRef buffer As Variant, ByVal index As Long) As String
    BufferKey = ScopeKey(GdtInvoiceIdentity(CStr(buffer(index, 0)), CStr(buffer(index, 3)), _
        CStr(buffer(index, 1)), CStr(buffer(index, 2))), IIf(buffer(index, 4) = 1, "query", "sco-query"))
End Function

Public Function CanResumeGdtStage(ByRef buffer As Variant, ByVal index As Long, ByVal stage As String, _
    Optional ByVal zipFolder As String = "", Optional ByVal extractedFolder As String = "") As Boolean
    Dim key As String, signatureKey As String, record As Variant, count As Long
    If Not GdtResumeEnabled Or mStages Is Nothing Then Exit Function
    signatureKey = BufferKey(buffer, index): key = signatureKey & KeyPart(stage)
    If stage = "XML" Then key = key & KeyPart(zipFolder) & KeyPart(extractedFolder)
    If Not mSignatures.Exists(signatureKey) Or Not mStages.Exists(key) Then Exit Function
    record = mStages(key)
    If CStr(record(0)) <> CStr(mSignatures(signatureKey)) Then Exit Function
    If stage = "DETAIL" Then
        If mDetailCounts.Exists(signatureKey) Then count = CLng(mDetailCounts(signatureKey))
        CanResumeGdtStage = (count > 0 And count = CLng(record(1)))
    ElseIf stage = "XML" Then
        Dim proof As String
        proof = XmlFileProof(buffer, index, zipFolder, extractedFolder)
        CanResumeGdtStage = (Len(proof) > 0 And proof = CStr(record(3)))
    End If
    If CanResumeGdtStage Then
        GdtStagesSkipped = GdtStagesSkipped + 1
        If GdtReportAccount = mAccount Then
            Dim direction As String, source As String
            direction = IIf(mDirection = "1", UniConvert("Mua vafo"), UniConvert("Basn ra"))
            source = IIf(buffer(index, 4) = 1, "query", "sco-query")
            MarkGdtRetrySuccess direction, source, CStr(buffer(index, 0)), CStr(buffer(index, 3)), CStr(buffer(index, 1)), _
                CStr(buffer(index, 2)), IIf(stage = "DETAIL", UniConvert("Laasy chi tieest"), UniConvert("Tari XML")), 0
            If stage = "DETAIL" Then MarkGdtRetrySuccess direction, source, CStr(buffer(index, 0)), CStr(buffer(index, 3)), _
                CStr(buffer(index, 1)), CStr(buffer(index, 2)), "Parse " & UniConvert("duwx lieeju"), 0
        End If
    End If
End Function

Private Function XmlFileProof(ByRef buffer As Variant, ByVal index As Long, ByVal zipFolder As String, ByVal extractedFolder As String) As String
    Dim fso As Object, shell As Object, archive As Object, files As New Collection, entry As Object
    Dim base As String, zipPath As String, path As String, extension As String, suffix As String
    Dim proof As Object, seen As Object, doc As Object
    On Error GoTo Missing
    Set fso = CreateObject("Scripting.FileSystemObject")
    base = GdtInvoiceFileBase(CStr(buffer(index, 0)), CStr(buffer(index, 3)), CStr(buffer(index, 1)), CStr(buffer(index, 2)))
    zipPath = fso.BuildPath(zipFolder, base & ".zip")
    If Not fso.FileExists(zipPath) Then Exit Function
    Set shell = CreateObject("Shell.Application")
    Set archive = shell.Namespace(CVar(fso.GetAbsolutePathName(zipPath)))
    If archive Is Nothing Then Exit Function
    CollectGdtZipFiles archive, files, 0
    If files.Count = 0 Then Exit Function
    Set proof = CreateObject("Scripting.Dictionary"): Set seen = CreateObject("Scripting.Dictionary")
    proof("zip") = CDbl(fso.GetFile(zipPath).Size)
    For Each entry In files
        extension = LCase$(fso.GetExtensionName(CStr(entry.Name)))
        suffix = ""
        If seen.Exists(extension) Then
            seen(extension) = CLng(seen(extension)) + 1: suffix = "_" & seen(extension)
        Else
            seen(extension) = 1
        End If
        path = fso.BuildPath(extractedFolder, base & suffix & "." & extension)
        If Not fso.FileExists(path) Then Exit Function
        If CDbl(fso.GetFile(path).Size) <> CDbl(entry.Size) Then Exit Function
        If extension = "xml" Then
            Set doc = CreateObject("MSXML2.DOMDocument.6.0")
            doc.async = False: doc.resolveExternals = False: doc.setProperty "ProhibitDTD", True
            If Not doc.Load(path) Then Exit Function
        End If
        proof(base & suffix & "." & extension) = CDbl(entry.Size)
    Next entry
    XmlFileProof = JsonConverter.ConvertToJson(proof)
Missing:
End Function

Public Sub MarkGdtStageComplete(ByRef buffer As Variant, ByVal index As Long, ByVal stage As String, _
    ByVal rowCount As Long, Optional ByVal zipFolder As String = "", Optional ByVal extractedFolder As String = "")
    Dim key As String, signatureKey As String, ws As Worksheet, targetRow As Long, record As Variant, proof As String
    If mStages Is Nothing Then Exit Sub
    signatureKey = BufferKey(buffer, index)
    If Not mSignatures.Exists(signatureKey) Then Exit Sub
    If Len(CStr(mSignatures(signatureKey))) > 32767 Then Exit Sub
    key = signatureKey & KeyPart(stage)
    If stage = "XML" Then key = key & KeyPart(zipFolder) & KeyPart(extractedFolder)
    If stage = "XML" Then
        proof = XmlFileProof(buffer, index, zipFolder, extractedFolder)
        If Len(proof) = 0 Or Len(proof) > 32767 Then Exit Sub
    End If
    Set ws = GetStateSheet(True)
    If mStages.Exists(key) Then
        record = mStages(key): targetRow = CLng(record(2))
    Else
        targetRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    End If
    GdtWriteTextCell ws.Cells(targetRow, 1), key
    GdtWriteTextCell ws.Cells(targetRow, 2), CStr(mSignatures(signatureKey))
    ws.Cells(targetRow, 3).Value2 = stage: ws.Cells(targetRow, 4).Value2 = rowCount
    ws.Cells(targetRow, 5).Value2 = CDbl(Now)
    GdtWriteTextCell ws.Cells(targetRow, 6), proof
    mStages(key) = Array(CStr(mSignatures(signatureKey)), rowCount, targetRow, proof)
End Sub

Public Sub CommitGdtDetailRows(ByRef buffer As Variant, ByVal index As Long, ByVal firstNewRow As Long, ByRef nextRow As Long)
    Dim ws As Worksheet, key As String, source As String, count As Long, identity As String
    If mStages Is Nothing Then Exit Sub
    Set ws = ThisWorkbook.Sheets(IIf(mDirection = "1", "ChiTietHD_Mua", "ChiTietHD_Ban"))
    count = nextRow - firstNewRow
    key = BufferKey(buffer, index)
    source = IIf(buffer(index, 4) = 1, "query", "sco-query")
    If mDetailCounts.Exists(key) Then mReplaced(key) = True
    identity = GdtInvoiceIdentity(CStr(buffer(index, 0)), CStr(buffer(index, 3)), CStr(buffer(index, 1)), CStr(buffer(index, 2)))
    If mLegacyDetails.Exists(identity) Then mLegacyReplaced(identity) = True
    ws.Cells(2, 35).Value2 = "Account scope": ws.Cells(2, 36).Value2 = "API source"
    ws.Columns(35).Hidden = True: ws.Columns(36).Hidden = True
    ws.Cells(firstNewRow, 35).Resize(count, 1).NumberFormat = "@"
    ws.Cells(firstNewRow, 35).Resize(count, 1).Value2 = mAccount
    ws.Cells(firstNewRow, 36).Resize(count, 1).NumberFormat = "@"
    ws.Cells(firstNewRow, 36).Resize(count, 1).Value2 = source
    mDetailCounts(key) = count
    MarkGdtStageComplete buffer, index, "DETAIL", count
End Sub

Public Sub FinalizeGdtResumeDetails()
    Dim ws As Worksheet, data As Variant, r As Long, key As String, groupEnd As Long, remove As Boolean
    If mReplaced Is Nothing Then Exit Sub
    If (mReplaced.Count = 0 And mLegacyReplaced.Count = 0) Or mOriginalDetailLast < 3 Then Exit Sub
    Set ws = ThisWorkbook.Sheets(IIf(mDirection = "1", "ChiTietHD_Mua", "ChiTietHD_Ban"))
    data = ws.Cells(3, 1).Resize(mOriginalDetailLast - 2, 36).Value2
    ' Delete old owned blocks once, bottom-up, only after replacement writes succeeded.
    For r = UBound(data, 1) To 1 Step -1
        remove = False
        If CStr(data(r, 35)) = mAccount Then
            key = ScopeKey(GdtInvoiceIdentity(CStr(data(r, 8)), CStr(data(r, 1)), CStr(data(r, 2)), CStr(data(r, 3))), CStr(data(r, 36)))
            remove = mReplaced.Exists(key)
        ElseIf Len(CStr(data(r, 35))) = 0 And Len(mAccount) > 0 Then
            If CStr(data(r, IIf(mDirection = "1", 14, 8))) = mAccount Then
                key = GdtInvoiceIdentity(CStr(data(r, 8)), CStr(data(r, 1)), CStr(data(r, 2)), CStr(data(r, 3)))
                remove = mLegacyReplaced.Exists(key)
            End If
        End If
        If remove Then
            If groupEnd = 0 Then groupEnd = r + 2
        ElseIf groupEnd > 0 Then
            ws.Rows(CStr(r + 3) & ":" & groupEnd).Delete
            groupEnd = 0
        End If
    Next r
    If groupEnd > 0 Then ws.Rows("3:" & groupEnd).Delete
    mReplaced.RemoveAll
    mLegacyReplaced.RemoveAll
End Sub
