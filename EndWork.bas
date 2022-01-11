Attribute VB_Name = "Module4"
Sub SendEmail_GoodBye()

Dim objOutlook As Outlook.Application
Dim objMail As Outlook.MailItem
Dim wsMail As Worksheet

Set objOutlook = New Outlook.Application
Set wsMail = ThisWorkbook.Sheets("メール内容")

Set objMail = objOutlook.CreateItem(olMailItem)

With wsMail

    objMail.To = .Range("A9").Value        'メール宛先
    objMail.Subject = .Range("D2").Value   'メール件名
    objMail.BodyFormat = olFormatPlain     'メールの形式
    objMail.Body = .Range("D3").Value      'メール本文

    objMail.Send
End With

Set objOutlook = Nothing
MsgBox "送信完了(業務終了)"

End Sub


