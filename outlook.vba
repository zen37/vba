Option Explicit
 
Private WithEvents oItems As Outlook.Items
 
Private Const FILE_PATH As String = "...." 'file path and name to save some info from the deleted reply
 
Private Const REPLY_LENGTH As Integer = 99 'minimum reply size to be considered "valuable"
 
Private Sub Application_Startup()
 
Dim oNS  As Outlook.NameSpace
 
Dim item As Object
Dim i As Integer
 
    Set oNS = Outlook.Application.GetNamespace("MAPI")
 
    Set oItems = oNS.GetDefaultFolder(olFolderInbox).Items
       
    If Not (oItems Is Nothing) Then
        
     '   For Each item In oItems
      For i = oItems.Count To 1 Step -1
     
      Set item = oItems.item(i)
            If item.Class = olMail Then
                MoveItem item
            End If
       Next
 
    End If
 
End Sub
 
Private Sub oItems_ItemAdd(ByVal item As Object)
 
    Dim oMail As Outlook.MailItem
    Dim oConv As Outlook.Conversationa
   
    Dim Pos As Long
    Dim LengthReply As Integer
    Dim reply As String
 
    If TypeName(item) = "MailItem" Then
       
         MoveItem item
     
    End If
 
End Sub
 
Private Sub MoveItem(item As Object)
 
    Dim oMail As Outlook.MailItem
    Dim oConv As Outlook.Conversation
    Dim Pos As Long
    Dim LengthReply As Integer
    Dim reply As String

    Set oMail = item
    Set oConv = oMail.GetConversation
 
    If Not (oConv Is Nothing) Then
        
      If Len(oMail.ConversationIndex) > 44 Then 'this is a Reply
 
        Pos = InStr(oMail.Body, "From:")
       
        If Pos >= 1 Then
      
         reply = Left(oMail.Body, Pos - 1)
                
         If reply <> "" Then
        
            LengthReply = Len(reply)
           
       '     Debug.Print "reply length: " & LengthReply & " subject: " & oMail.subject
                   
            If LengthReply < REPLY_LENGTH Then
   
                File_Write oMail.ReceivedTime, oMail.subject, reply
                oMail.Delete
                      
            End If
           
         End If
               
        End If
        
      End If
       
    End If
 
End Sub
 
 
Private Sub File_Write(received As String, subject As String, reply As String)
 
'PURPOSE: Add More Text To The End Of A Text File
'SOURCE: www.TheSpreadsheetGuru.com
 
Dim TextFile As Integer
Dim FilePath As String
 
'What is the file path and name for the new text file?
  FilePath = FILE_PATH
 
'Determine the next file number available for use by the FileOpen function
  TextFile = FreeFile
 
'Open the text file
  Open FilePath For Append As TextFile
 
'Write some lines of text
  Print #TextFile, received & " subject: " & subject
  Print #TextFile, RemoveObsoleteWhiteSpace(reply)
  Print #TextFile, "------------------------------------------"
 
'Save & Close Text File
  Close TextFile
End Sub
 
Public Function RemoveObsoleteWhiteSpace(FromString As Variant) As Variant
 
'https://social.msdn.microsoft.com/Forums/office/en-US/2f3f9a60-4da6-40a2-828a-8c9b58586119/is-there-a-way-to-remove-all-extra-spaces-in-a-string
 
  If IsNull(FromString) Then 'handle Null values
    RemoveObsoleteWhiteSpace = Null
    Exit Function
  End If
 
  Dim strTemp As String
 
  strTemp = Replace(FromString, vbCr, " ")
  strTemp = Replace(strTemp, vbLf, " ")
  strTemp = Replace(strTemp, vbTab, " ")
  strTemp = Replace(strTemp, vbVerticalTab, " ")
  strTemp = Replace(strTemp, vbBack, " “)
  strTemp = Replace(strTemp, vbNullChar, " ")
 
  While InStr(strTemp, "  ") > 0
    strTemp = Replace(strTemp, "  ", " ")
  Wend
 
  strTemp = Trim(strTemp)
  RemoveObsoleteWhiteSpace = strTemp
 
End Function