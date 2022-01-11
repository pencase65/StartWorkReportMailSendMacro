Attribute VB_Name = "Module1"
Sub SendEmail_Hello()

Dim objOutlook As Outlook.Application
Dim objMail As Outlook.MailItem
Dim wsMail As Worksheet

Set objOutlook = New Outlook.Application
Set wsMail = ThisWorkbook.Sheets("メール内容")

Set objMail = objOutlook.CreateItem(olMailItem)

With wsMail

    objMail.To = .Range("A9").Value        'メール宛先
    objMail.Subject = .Range("A2").Value   'メール件名
    objMail.BodyFormat = olFormatPlain     'メールの形式
    objMail.Body = .Range("A3").Value      'メール本文

    objMail.Send
End With

Set objOutlook = Nothing
MsgBox "送信完了(勤務開始)"

End Sub
