Attribute VB_Name = "Module1"
Sub SendEmail_Hello()

Dim objOutlook As Outlook.Application
Dim objMail As Outlook.MailItem
Dim wsMail As Worksheet

Set objOutlook = New Outlook.Application
Set wsMail = ThisWorkbook.Sheets("���[�����e")

Set objMail = objOutlook.CreateItem(olMailItem)

With wsMail

    objMail.To = .Range("A9").Value        '���[������
    objMail.Subject = .Range("A2").Value   '���[������
    objMail.BodyFormat = olFormatPlain     '���[���̌`��
    objMail.Body = .Range("A3").Value      '���[���{��

    objMail.Send
End With

Set objOutlook = Nothing
MsgBox "���M����(�Ζ��J�n)"

End Sub
