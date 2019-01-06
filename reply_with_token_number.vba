Option Explicit
Dim token As Double


Sub AutoReply(olItem As Outlook.MailItem)
    Dim olOutMail As Outlook.MailItem
    
'generate token
   Dim myfile As String, textline As String, num As Double, numfile As String
   
   myfile = "C:\Users\MASBHA\Documents\auto_num.txt"
   numfile = FreeFile
   
   Open myfile For Input As numfile
   Do Until EOF(numfile)
       Line Input #1, textline
       num = textline
   Loop
   Close numfile
   
   Dim auto_num As Double
   auto_num = num + 1
   
   Open myfile For Output As numfile
   Print #1, auto_num
   Close numfile
   token = auto_num
'end token

    With olItem
        Set olOutMail = olItem.Reply
        With olOutMail
            .Body = "Hi," & vbNewLine & vbNewLine & _
                    "SA number against your request is = " & token & vbNewLine & _
                    "Have a good day !" & vbNewLine & vbNewLine & _
                    "Thanks," & vbNewLine & _
                    "CSD"
            .Send
        End With
        Set olOutMail = Nothing
    End With
End Sub