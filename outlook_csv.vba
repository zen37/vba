Private Const FILE_PATH As String = "C:\Users\....."

 

Private Sub Application_Startup()

 

    ListMailsInFolder

 

 

End Sub

 

Private Sub ListMailsInFolder()

 

    Dim objNS As Outlook.NameSpace

    Dim objFolder As Outlook.MAPIFolder

   

    Dim inv() As String, file_name() As String

    Dim subject As String

   

    

 

    Set objNS = GetNamespace("MAPI")

    Set objFolder = objNS.Folders.GetFirst ' folders of your current account

    Set objFolder = objFolder.Folders("...").Folders("...").Folders("...").Folders("...")

 

    For Each item In objFolder.Items

        If TypeName(item) = "MailItem" Then

            ' ... do stuff here ...

           ' Debug.Print item.ConversationTopic & " - " & item.SentOn


            subject = item.ConversationTopic

          

            inv = Split(subject, "_")

           

            file_name = Split(subject)

           

            

            'Debug.Print inv(0) & " - " & file_name(0)

          

            File_Write inv(0), file_name(0), item.SentOn

           

        End If

    Next

 

End Sub

 

Private Sub File_Write2(i As String, f As String, sent_on As Date)

 

Dim fso As Object

Set fso = CreateObject("Scripting.FileSystemObject")

Dim oFile As Object

Set oFile = fso.CreateTextFile(FILE_PATH)

oFile.WriteLine i & "," & f & "," & sent_on

oFile.Close

Set fso = Nothing

Set oFile = Nothing

 

End Sub

 

 

 

Private Sub File_Write(i As String, f As String, sent_on As Date)

 

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

  Print #TextFile, i & "," & f & "," & sent_on

'Print #TextFile, RemoveObsoleteWhiteSpace(reply)

' Print #TextFile, "------------------------------------------"

 

'Save & Close Text File

  Close TextFile

End Sub