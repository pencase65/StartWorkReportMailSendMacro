Attribute VB_Name = "Module4"
Sub SendEmail_GoodBye()

Dim objOutlook As Outlook.Application
Dim objMail As Outlook.MailItem
Dim wsMail As Worksheet

Set objOutlook = New Outlook.Application
Set wsMail = ThisWorkbook.Sheets("���[�����e")

Set objMail = objOutlook.CreateItem(olMailItem)

With wsMail

    objMail.To = .Range("A9").Value        '���[������
    objMail.Subject = .Range("D2").Value   '���[������
    objMail.BodyFormat = olFormatPlain     '���[���̌`��
    objMail.Body = .Range("D3").Value      '���[���{��

    objMail.Send
End With

Set objOutlook = Nothing
MsgBox "���M����(�Ɩ��I��)"

End Sub


