Attribute VB_Name = "SaveCoordination"
Option Explicit

Sub SaveAttachmentsToFolder()
    Dim objMail As MailItem
    Dim objAttachments As Attachments
    Dim objAttachment As Attachment
    Dim strProjectNumber As String
    Dim strFolderName As String
    Dim strFolderPath As String
    Dim strDate As String
    Dim fso As Object

    ' Prompt for project number
    strProjectNumber = InputBox("Enter Project Number:", "Project Number")
    If strProjectNumber = "" Then Exit Sub
    
    ' Prompt for folder name
    strFolderName = InputBox("Enter Folder Name:", "Folder Name")
    If strFolderName = "" Then Exit Sub

    ' Get today's date in YYYY-MM-DD format
    strDate = Format(Date, "yyyy-mm-dd")

    ' Create folder name with date
    strFolderName = strDate & " - " & strFolderName

    ' Create folder path
    strFolderPath = "P:\" & strProjectNumber & "\Coordination\Received\" & strFolderName

    ' Create the folder if it does not exist
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(strFolderPath) Then
        fso.CreateFolder strFolderPath
    End If

    ' Check if an email is selected
    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "Please select an email first.", vbExclamation
        Exit Sub
    End If

    ' Get the selected email
    Set objMail = Application.ActiveExplorer.Selection.Item(1)

    ' Get the attachments collection
    Set objAttachments = objMail.Attachments

    ' Save each attachment to the specified folder
    For Each objAttachment In objAttachments
        objAttachment.SaveAsFile strFolderPath & "\" & objAttachment.FileName
    Next objAttachment

    ' Copy the folder path to the clipboard
    CopyTextToClipboard strFolderPath

    ' Notify the user
    MsgBox "Attachments saved to: " & strFolderPath, vbInformation
End Sub

Sub CopyTextToClipboard(ByVal strText As String)
    Dim objData As MSForms.DataObject
    Set objData = New MSForms.DataObject
    objData.SetText strText
    objData.PutInClipboard
End Sub

