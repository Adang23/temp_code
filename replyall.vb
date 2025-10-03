Option Explicit

Public Sub PrepareReplyEmail()
    On Error GoTo CleanFail

    Dim src As MailItem
    Set src = GetCurrentMailItem()
    If src Is Nothing Then Exit Sub

    ' Always Reply All
    Dim r As MailItem
    Set r = src.ReplyAll

    ' Keep reply header, then restore original HTML body
    Dim replyHeader As String
    replyHeader = r.HTMLBody
    r.HTMLBody = replyHeader & src.HTMLBody

    ' Temp folder
    Dim fso As Object, tempPath As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    tempPath = Environ$("TEMP") & "\OutlookReplyTmp\"
    If Not fso.FolderExists(tempPath) Then fso.CreateFolder tempPath

    ' Copy attachments (preserve Content-ID for inline images)
    Dim a As Attachment, newA As Attachment
    Dim tmpFile As String, cid As String
    Dim pa As PropertyAccessor

    For Each a In src.Attachments
        On Error Resume Next
        tmpFile = tempPath & a.FileName
        a.SaveAsFile tmpFile
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            GoTo NextAttachment
        End If
        On Error GoTo 0

        Set newA = r.Attachments.Add(tmpFile)

        ' Preserve inline rendering
        On Error Resume Next
        cid = a.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")
        On Error GoTo 0
        If Len(cid) > 0 Then
            Set pa = newA.PropertyAccessor
            On Error Resume Next
            pa.SetProperty "http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid
            pa.SetProperty "http://schemas.microsoft.com/mapi/proptag/0x37140003", 1
            On Error GoTo 0
        End If
NextAttachment:
    Next a

    ' Just show draft, do not send
    r.Display
    Exit Sub

CleanFail:
    MsgBox "Couldn't prepare reply email: " & Err.Description, vbExclamation
End Sub

Private Function GetCurrentMailItem() As MailItem
    On Error Resume Next
    If Not Application.ActiveInspector Is Nothing Then
        If TypeOf Application.ActiveInspector.CurrentItem Is MailItem Then
            Set GetCurrentMailItem = Application.ActiveInspector.CurrentItem
            Exit Function
        End If
    End If
    If Not Application.ActiveExplorer Is Nothing Then
        Dim sel As Selection
        Set sel = Application.ActiveExplorer.Selection
        If Not sel Is Nothing Then
            If sel.Count > 0 Then
                If TypeOf sel.Item(1) Is MailItem Then
                    Set GetCurrentMailItem = sel.Item(1)
                    Exit Function
                End If
            End If
        End If
    End If
    Set GetCurrentMailItem = Nothing
End Function
