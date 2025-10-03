' ReplyAll_PrepareWithImages.bas
' Prepares a Reply All that preserves original HTML (so inline images render)
' and re-attaches all attachments (incl. inline). It ONLY displays the draft.

Option Explicit

Public Sub ReplyAll_PrepareWithImages()
    PrepareReply keepAllRecipients:=True
End Sub

Public Sub Reply_PrepareWithImages()
    PrepareReply keepAllRecipients:=False
End Sub

Private Sub PrepareReply(ByVal keepAllRecipients As Boolean)
    On Error GoTo CleanFail

    Dim src As MailItem
    Set src = GetCurrentMailItem()
    If src Is Nothing Then Exit Sub

    ' Create reply (no send!)
    Dim r As MailItem
    If keepAllRecipients Then
        Set r = src.ReplyAll
    Else
        Set r = src.Reply
    End If

    ' Keep Outlook's reply header (From/To/Date...) but replace the quoted body
    ' with the original HTML so <img src="cid:..."> works reliably.
    Dim replyHeader As String
    replyHeader = r.HTMLBody
    r.HTMLBody = replyHeader & src.HTMLBody

    ' Temp folder for copying attachments
    Dim fso As Object, tempPath As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    tempPath = Environ$("TEMP") & "\OutlookReplyTmp\"
    If Not fso.FolderExists(tempPath) Then fso.CreateFolder tempPath

    ' Re-attach everything; preserve Content-ID for inline images.
    Dim a As Attachment, newA As Attachment
    Dim tmpFile As String, cid As String
    Dim pa As PropertyAccessor

    For Each a In src.Attachments
        ' Skip cloud/placeholder attachments that can't be saved
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

        ' Copy PR_ATTACH_CONTENT_ID if present (renders inline)
        On Error Resume Next
        cid = a.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")
        On Error GoTo 0
        If Len(cid) > 0 Then
            Set pa = newA.PropertyAccessor
            On Error Resume Next
            pa.SetProperty "http://schemas.microsoft.com/mapi/proptag/0x3712001F", cid     ' Content-ID
            pa.SetProperty "http://schemas.microsoft.com/mapi/proptag/0x37140003", 1       ' Attach flags = inline
            On Error GoTo 0
        End If
NextAttachment:
    Next a

    ' Show the draft for manual editing/sending
    r.Display
    Exit Sub

CleanFail:
    MsgBox "Couldn't prepare reply with images/attachments: " & Err.Description, vbExclamation
End Sub

' Helper: pick the active mail either from an open window or the selection
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
