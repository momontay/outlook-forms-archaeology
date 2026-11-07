'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 2011
'
' NAME: Пиво Вода Akt_Spis_08e1v4
'
' AUTHOR: Mrs , 
' DATE  : 09.06.2016
'
' COMMENT: 
'
'==========================================================================

Function Item_Open()

Dim strUserDN, objSysInfo, objUser, userName
Dim sTargetFolder, ofso1, oFile, num   'Переменные для номеров заявок
Dim  sGroup
set WshShell=CreateObject("WScript.Shell") 
'----------------------------------из AD--------------------------------------------------
Set objSysInfo=CreateObject("ADSystemInfo") 
strUserDN=objSysInfo.userName 
Set objUser = GetObject("LDAP://" & strUserDN)
'---------------------------------------------------------------------------------------------
Set ofso1 = CreateObject("Scripting.FileSystemObject") 'Объект файловой системы
sTargetFolder = "\\your-file-server\systemf$\Number.txt"   'Файл с номерами
sTargetFolder1 = "\\your-file-server\systemf$\NumberUr.txt" 'Файл с номерами юристы
myFilePath = "\\your-file-server\systemf$\users.ini"

VarX=objUser.sAMAccountName

userName = objUser.sn+" "+objUser.givenName
'msgbox(userName)
Item.UserProperties.Find("app-userName") = userName
sGroup = ReadIni( myFilePath, VarX )
'msgbox(sGroup)
Item.UserProperties.Find("app-buffgroup") = sGroup

'----------------------------------------Права доступа к форме---------------------------------------------------------------------------------------

If (sGroup = "admins") Then 

   ' Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("ComboBox4").Visible = True
    'Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("ButtonPometka9").Visible = True
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox36").Locked = False
    'Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox38").Locked = False
    'Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox38").Enabled = True
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame1").Enabled = True
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame2").Enabled = True
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame3").Enabled = True
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame4").Enabled = True
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame5").Enabled = True   'Администраторы
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame6").Enabled = True
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame7").Enabled = True
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame8").Enabled = True 
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame9").Enabled = True     
ElseIf  (sGroup = "torgpr")  Then
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame1").Enabled = True
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame2").Enabled = True
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame3").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame4").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame5").Enabled = False   'Торговые представители
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame6").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame7").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame8").Enabled = False 
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame9").Enabled = False 
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame2").BorderStyle = 1
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame2").BorderColor = RGB(255, 0, 0)    
ElseIf  (sGroup = "superv")  then
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame1").Enabled = True
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame2").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame3").Enabled = True
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame4").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame5").Enabled = False    'Супервайзеры
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame6").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame7").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame8").Enabled = False 
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame9").Enabled = False 
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame3").BorderStyle = 1
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame3").BorderColor = RGB(255, 0, 0)      
ElseIf  (sGroup = "rukprod")  Then
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame1").Enabled = True
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame2").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame3").Enabled = True
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame4").Enabled = True
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame5").Enabled = False    'Руководитель отдела продаж РОП
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame6").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame7").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame8").Enabled = False 
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame9").Enabled = False  
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame4").BorderStyle = 1
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame4").BorderColor = RGB(255, 0, 0)
ElseIf  (sGroup = "komdir") Then
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame1").Enabled = True
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame2").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame3").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame4").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame5").Enabled = False    'Комерческий директор
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame6").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame7").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame8").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame9").Enabled = True
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame9").BorderStyle = 1
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame9").BorderColor = RGB(255, 0, 0)
ElseIf  (sGroup = "kreditk")  Then
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame1").Enabled = True
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame2").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame3").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame4").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame5").Enabled = True     'Кредитный контролер
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame6").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame7").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame8").Enabled = False 
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame9").Enabled = False  
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame5").BorderStyle = 1
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame5").BorderColor = RGB(255, 0, 0)   
    '------------------
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("ComboBox1").Enabled = False  
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox1").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox2").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox3").Enabled = False 
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox4").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox5").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox6").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox7").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox8").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox9").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox10").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox11").Enabled = False  
ElseIf  (sGroup = "urotdel")  Then       '(Item.UserProperties.Find("app-CheckBox3")=True Or And  Item.UserProperties.Find("app-toSB")="Yes"
       Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame1").Enabled = True
	   Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame2").Enabled = False
	   Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame3").Enabled = False
	   Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame4").Enabled = True
	   Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame5").Enabled = False   'Юристы
	   Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame6").Enabled = True
	   Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame7").Enabled = False
       Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame8").Enabled = False 
       Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame9").Enabled = False  
       Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame6").BorderStyle = 1
       Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame6").BorderColor = RGB(255, 0, 0)
       'Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("CommandButton4").Enabled = True        
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("ComboBox1").Enabled = False  
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox1").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox2").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox3").Enabled = False 
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox4").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox5").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox6").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox7").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox8").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox9").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox10").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox11").Enabled = False
ElseIf  (sGroup = "findir")  Then
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame1").Enabled = True
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame2").Enabled = True
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame3").Enabled = True
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame4").Enabled = True
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame5").Enabled = True    'Финансовый директор
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame6").Enabled = True
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame7").Enabled = True  
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame8").Enabled = True 
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame9").Enabled = True 
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame7").BorderStyle = 1
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame7").BorderColor = RGB(255, 0, 0)
    'Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame7").BorderColor = RGB(0, 0, 255)     
ElseIf  (sGroup = "fclose") And Item.UserProperties.Find("app-CheckBox7")=True then
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame1").Enabled = True
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame2").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame3").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame4").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame5").Enabled = False    'Закрывающий
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame6").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame7").Enabled = False   
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame8").Enabled = True    
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame9").Enabled = False  
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame8").BorderStyle = 1
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame8").BorderColor = RGB(255, 0, 0) 
    'Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame8").BorderColor = RGB(0, 0, 255)  
    '-----Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("CommandButton4").Enabled = True 
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("ComboBox1").Enabled = False   
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox1").Enabled = False  
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox2").Enabled = False  
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox3").Enabled = False  
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox4").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox5").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox6").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox7").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox8").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox9").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox10").Enabled = False
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox11").Enabled = False
Else
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame1").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame2").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame3").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame4").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame5").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame6").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame7").Enabled = False   
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame8").Enabled = False  
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame9").Enabled = False      
         
End If

'-------------------------------Удалена не удалена-------------------------------------------------------------------------------------------------------------------
If Item.UserProperties.Find("Статус заявки")="Удалено" Then
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame1").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame2").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame3").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame4").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame5").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame6").Enabled = False
	Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame7").Enabled = False   
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame8").Enabled = False  
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame9").Enabled = False   
End If
'-------------------Опции----------------------------------------------------------------

If Item.UserProperties.Find("app-optbtnbuff")="OptionButton4" Then
  Set OptionButton4 = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton4")
  OptionButton4.Value=True
ElseIf Item.UserProperties.Find("app-optbtnbuff")="OptionButton5" Then
  Set OptionButton5 = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton5")
  OptionButton5.Value=True
ElseIf Item.UserProperties.Find("app-optbtnbuff")="OptionButton6" Then
  Set OptionButton6 = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton6")
  OptionButton6.Value=True  
End If 
'------------------Опции2-----------------------------------------------------------------

If Item.UserProperties.Find("app-optbtnbuff1")="OptionButton1" Then
  Set OptionButton1 = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton1")
  OptionButton1.Value=True
ElseIf Item.UserProperties.Find("app-optbtnbuff1")="OptionButton2" Then
  Set OptionButton2 = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton2")
  OptionButton2.Value=True
ElseIf Item.UserProperties.Find("app-optbtnbuff1")="OptionButton3" Then
  Set OptionButton3 = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton3")
  OptionButton3.Value=True  
End If 
'-----------------Опции3/4-----------------------------------------------------------------

If Item.UserProperties.Find("app-optbtnbuff2")="OptionButton7" Then
  Set OptionButton1 = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton7")
  OptionButton1.Value=True
ElseIf Item.UserProperties.Find("app-optbtnbuff2")="OptionButton8" Then
  Set OptionButton2 = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton8")
  OptionButton2.Value=True 
End If 
'------------
If Item.UserProperties.Find("app-optbtnbuff3")="OptionButton9" Then
  Set OptionButton1 = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton9")
  OptionButton1.Value=True
ElseIf Item.UserProperties.Find("app-optbtnbuff3")="OptionButton10" Then
  Set OptionButton2 = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton10")
  OptionButton2.Value=True
End If 
'------------


  If Item.UserProperties.Find("app-datepostupl")<>"" Then 
     Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("CommandButton2").Enabled = False
     Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame1").Enabled = False
     Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame2").Enabled = False
  	 Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame3").Enabled = False
 	 Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame4").Enabled = False
 	 'Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame5").Enabled = False
  End If

 If Item.UserProperties.Find("app-saved")= "Saved" Then Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("CommandButton4").Enabled = False
 'If Item.UserProperties.Find("app-saved")= "Saved" Then 
    'Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("CommandButton8").Enabled = True
' Else 
    'Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("CommandButton8").Enabled = False
 'End If

End Function

'**********************************************************************************************************************************************************

Sub SendMail(mails)

sLink = "<a href='outlook: " + item.EntryID + " '> нажмите здесь </ A>"
strSmtpServer="your-email-server" ' 
strSmtpPort=25
strSmtpAuth="no"
strSmtpSsl="no"
strSendingEmail="notify@domain.ru" ' Имя отправителя
strReportEmail=mails ' Имя получателя
Set objMessage = CreateObject("CDO.Message")
objMessage.BodyPart.CharSet="utf-8"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSmtpServer
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = strSmtpPort
If strSmtpAuth = "yes" Then
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'use '2' for NTLM authentication
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = strSmtpUser
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strSmtpPass
End If
If strSmtpSsl = "yes" Then
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
End If
objMessage.Configuration.Fields.Update
objMessage.Subject = "Заявка на списание" + Item.UserProperties.Find("Номер заявки")
objMessage.From = strSendingEmail
objMessage.To = strReportEmail
objMessage.HTMLBody = "Заявка на списание дебиторской задолженности №" + Item.UserProperties.Find("Номер заявки") + "- ожидает вашего рассмотрения" + " - " + sLink
objMessage.Send

End Sub
'**************************************************************************************************************************************************
Sub SendMail1(mails)

sLink = "<a href='outlook: " + item.EntryID + " '> нажмите здесь </ A>"
strSmtpServer="your-email-server" ' 
strSmtpPort=25
strSmtpAuth="no"
strSmtpSsl="no"
strSendingEmail="notify@domain.ru" ' Имя отправителя
strReportEmail=mails ' Имя получателя
Set objMessage = CreateObject("CDO.Message")
objMessage.BodyPart.CharSet="utf-8"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSmtpServer
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = strSmtpPort
If strSmtpAuth = "yes" Then
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'use '2' for NTLM authentication
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = strSmtpUser
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strSmtpPass
End If
If strSmtpSsl = "yes" Then
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
End If
objMessage.Configuration.Fields.Update
objMessage.Subject = "Новая заявка на списание"
objMessage.From = strSendingEmail
objMessage.To = strReportEmail
objMessage.HTMLBody = "Появилась заявка на списание дебиторской задолженности №" + Item.UserProperties.Find("Номер заявки") + "  Юр. лицо клиента: - " + Item.UserProperties.Find("app-nazvclient") + " - " + sLink 
objMessage.Send

End Sub
'******************************************************************************************************************************************************
Sub SendMail2(mails, sText)

sLink = "<a href='outlook: " + item.EntryID + " '> нажмите здесь </ A>"
strSmtpServer="your-email-server" ' 
strSmtpPort=25
strSmtpAuth="no"
strSmtpSsl="no"
strSendingEmail="notify@domain.ru" ' Имя отправителя
strReportEmail=mails ' Имя получателя
Set objMessage = CreateObject("CDO.Message")
objMessage.BodyPart.CharSet="utf-8"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSmtpServer
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = strSmtpPort
If strSmtpAuth = "yes" Then
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'use '2' for NTLM authentication
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = strSmtpUser
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strSmtpPass
End If
If strSmtpSsl = "yes" Then
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
End If
objMessage.Configuration.Fields.Update
objMessage.Subject = "Заявка возвращена на пересмотр"
objMessage.From = strSendingEmail
objMessage.To = strReportEmail
objMessage.HTMLBody = sText + "  Заявка №" + Item.UserProperties.Find("Номер заявки") + "  Юр. лицо клиента: - " + Item.UserProperties.Find("app-nazvclient") + " - " + sLink 
objMessage.Send

End Sub
'***********************************************************************************************************************************************
Sub SendMail3(mails, sText)

sLink = "<a href='outlook: " + item.EntryID + " '> нажмите здесь </ A>"
strSmtpServer="your-email-server" ' имя почтового сервера, в моем случае это Exchange
strSmtpPort=25
strSmtpAuth="no"
strSmtpSsl="no"
strSendingEmail="notify@domain.ru" ' Имя отправителя
strReportEmail=mails ' Имя получателя
Set objMessage = CreateObject("CDO.Message")
objMessage.BodyPart.CharSet="utf-8"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSmtpServer
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = strSmtpPort
If strSmtpAuth = "yes" Then
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'use '2' for NTLM authentication
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = strSmtpUser
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strSmtpPass
End If
If strSmtpSsl = "yes" Then
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
End If
objMessage.Configuration.Fields.Update
objMessage.Subject = "Юридический отдел"
objMessage.From = strSendingEmail
objMessage.To = strReportEmail
objMessage.HTMLBody = sText + "  Заявка №" + Item.UserProperties.Find("Номер заявки") + "  Юр. лицо клиента: - " + Item.UserProperties.Find("app-nazvclient") + " - " + sLink 
objMessage.Send

End Sub
'***********************************************************************************************************************************************
Sub SendMail4(mails, sText, sFrom)

sLink = "<a href='outlook: " + item.EntryID + " '> нажмите здесь </ A>"
strSmtpServer="your-email-server" ' 
strSmtpPort=25
strSmtpAuth="no"
strSmtpSsl="no"
strSendingEmail=sFrom ' Имя отправителя
strReportEmail=mails ' Имя получателя
Set objMessage = CreateObject("CDO.Message")
objMessage.BodyPart.CharSet="utf-8"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSmtpServer
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = strSmtpPort
If strSmtpAuth = "yes" Then
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'use '2' for NTLM authentication
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = strSmtpUser
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strSmtpPass
End If
If strSmtpSsl = "yes" Then
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
End If
objMessage.Configuration.Fields.Update
objMessage.Subject = "Заявка возвращена на пересмотр"
objMessage.From = strSendingEmail
objMessage.To = strReportEmail
objMessage.HTMLBody = sText + "  Заявка №" + Item.UserProperties.Find("Номер заявки") + "  Юр. лицо клиента: - " + Item.UserProperties.Find("app-nazvclient") + " - " + sLink 
objMessage.Send

End Sub
'**********************************************************************************************************
Sub SendMail5(mails, sText)

sLink = "<a href='outlook: " + item.EntryID + " '> нажмите здесь </ A>"
strSmtpServer="your-email-server" ' 
strSmtpPort=25
strSmtpAuth="no"
strSmtpSsl="no"
strSendingEmail="notify@domain.ru" ' Имя отправителя
strReportEmail=mails ' Имя получателя
Set objMessage = CreateObject("CDO.Message")
objMessage.BodyPart.CharSet="utf-8"
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = strSmtpServer
objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = strSmtpPort
If strSmtpAuth = "yes" Then
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'use '2' for NTLM authentication
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = strSmtpUser
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = strSmtpPass
End If
If strSmtpSsl = "yes" Then
	objMessage.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
End If
objMessage.Configuration.Fields.Update
objMessage.Subject = "Повторно. Заявка на списание ожидает вашего рассмотрения"
objMessage.From = strSendingEmail
objMessage.To = strReportEmail
objMessage.HTMLBody = sText + "  Заявка №" + Item.UserProperties.Find("Номер заявки") + "  Юр. лицо клиента: - " + Item.UserProperties.Find("app-nazvclient") + " - " + sLink 
objMessage.Send

End Sub
'*********************************************************************************************************************************************************
Function ReadIni( myFilePath, VarX )
    
    Const ForReading   = 1
    Const ForWriting   = 2
    Const ForAppending = 8

    Dim intEqualPos
    Dim objFSO, objIniFile
    Dim strFilePath, strKey, strLeftString, strLine, strSection

    Set objFSO = CreateObject( "Scripting.FileSystemObject" )

    ReadIni     = ""
    strFilePath = Trim( myFilePath )
    strSection  = Trim( mySection )
    strKey      = Trim( myKey )

    If objFSO.FileExists( strFilePath ) Then
        Set objIniFile = objFSO.OpenTextFile( strFilePath, ForReading, False )
        Do While objIniFile.AtEndOfStream = False
            strLine = Trim( objIniFile.ReadLine )

            ' Check if section is found in the current line
            If InStr(1,strLine,VarX) > 0  Then              
                    intEqualPos = InStr( 1, strLine, "=", 1 )
                    If intEqualPos > 0 Then
                        strLeftString = Trim( Left( strLine, intEqualPos - 1 ) )                                                
                            ReadIni = strLeftString
                            Else
                            ReadIni = ""                                                                                             
                    End If

                    ' Abort if the end of the INI file is reached
                    If objIniFile.AtEndOfStream Then Exit Do

                    ' Continue with next line
                    strLine = Trim( objIniFile.ReadLine )
                'Loop
            Exit Do
            End If
        Loop
        
        objIniFile.Close
    Else
        WScript.Echo strFilePath & " doesn't exists. Exiting..."
        Wscript.Quit 1
    End If
End Function

'**********************************************************************************************************************************************************
'-----------------Обработка кнопок-------------------------

'-------------------Комерческий----------------------------
Sub CommandButton1_Click() 'Комерческий посылает в СБ
   If Item.UserProperties.Find("app-toSB")="New" Then 
       
       sUserSuper = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("ComboBoxSuper").Value
       sLdp = "LDAP://CN="+sUserSuper+",ou=staff,ou=moscow,dc=domain,dc=ru" 
       On Error Resume Next 
       Set objSuper = GetObject(sLdp)
       'mails = objRop.mail
       'Call SendMail5(mails, sText)
       sMailSuper = objSuper.mail
       MsgBox("Юристам отправлено оповещение")
       sUserName = Item.UserProperties.Find("app-userName")
       Item.UserProperties.Find("app-fiokomdir")= sUserName
       Item.UserProperties.Find("app-datekd")=Cstr(Date())+"-"+Cstr(Time())
       Item.UserProperties.Find("Стадия")="Юридический отдел"
       Item.UserProperties.Find("app-toSB")="Saved"
       '----------------------------------------
          sUser1 = Item.UserProperties.Find("app-torgpredstavfio")
          If sUser1 = "" Then sUser1 = "PB"          
          sLdp1 = "LDAP://CN="+sUser1+",ou=staff,ou=moscow,dc=domain,dc=ru"
          Set objU1 = GetObject(sLdp1)
          sMail1 = objU1.mail 'Мыло манагера
          sMail2 = sMailSuper 'Мыло супервайзера
          mails02 = sMail1+";"+sMail2 
       sText = "Заявка отправлена в работу юридического отдела. Вам необходимо предоставить документы в юридический отдел в течение 10 календарных дней с момента получения этого письма!!! В случае непоступления документов заявка будет автоматически аннулирована!"
       Call SendMail3(mails02, sText)
       '----------------------------------------       
   End If
   On Error Resume Next
   Item.Save
End Sub

Sub CommandButtonSendAD_Click() 'Комерческий посылает в ЮА
     If Item.UserProperties.Find("app-toAD")="New" Then 
       mails01 = "name@domain.ru"
       Call SendMail(mails01)
       sUserSuper = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("ComboBoxSuper").Value
       sLdp = "LDAP://CN="+sUserSuper+",ou=staff,ou=moscow,dc=domain,dc=ru" 
       On Error Resume Next 
       Set objSuper = GetObject(sLdp)
       'mails = objRop.mail
       'Call SendMail5(mails, sText)
       sMailSuper = objSuper.mail
       MsgBox("В отдел СБ отправлено оповещение")
       sUserName = Item.UserProperties.Find("app-userName")
       Item.UserProperties.Find("app-rukprodaj")= sUserName
       Item.UserProperties.Find("app-daterukprod")=Cstr(Date())+"-"+Cstr(Time())
       Item.UserProperties.Find("Стадия")="Н Ю.А."
       Item.UserProperties.Find("app-toAD")="Saved"
       '----------------------------------------
          sUser1 = Item.UserProperties.Find("app-torgpredstavfio")
          If sUser1 = "" Then sUser1 = "PB"          
          sLdp1 = "LDAP://CN="+sUser1+",ou=staff,ou=moscow,dc=domain,dc=ru"
          Set objU1 = GetObject(sLdp1)
          sMail1 = objU1.mail 'Мыло манагера
          sMail2 = sMailSuper 'Мыло супервайзера
          mails02 = sMail1+";"+sMail2 
       sText = "Заявка отправлена в работу СБ (Н Ю.А.). Вам необходимо предоставить документы Ни Ю.А. в течение 10 календарных дней с момента получения этого письма!!! В случае непоступления документов заявка будет автоматически аннулирована!"
       Call SendMail3(mails02, sText)
       'mails = "alapskiy@alephtrade.com"
       'Call SendMail(mails)
       '----------------------------------------       
   End If
   On Error Resume Next
   Item.Save

End Sub

Sub CommandButtonSpisKd_Click()

If Item.UserProperties.Find("app-savedkd")="New" Then
   Set OptionButton4 = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton4")
   Set OptionButton5 = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton5")
   Set OptionButton6 = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton6")
   If OptionButton4.Value=True Then
      mails = "name@domain.ru"
      Call SendMail(mails)
      MsgBox("Финансовому директору отправлено оповещение")
      sUserName = Item.UserProperties.Find("app-userName")
      Item.UserProperties.Find("app-fiokomdir")= sUserName
      Item.UserProperties.Find("app-datekd")=Cstr(Date())+"-"+Cstr(Time())
      Item.UserProperties.Find("app-optbtnbuff")="OptionButton4"
      Item.UserProperties.Find("Стадия")="Финансовый директор"
      Item.UserProperties.Find("app-savedkd")="Saved"      
   ElseIf OptionButton5.Value=True Then
      mails = "name@domain.ru"
      Call SendMail(mails)
      MsgBox("Финансовому директору отправлено оповещение")
      sUserName = Item.UserProperties.Find("app-userName")
      Item.UserProperties.Find("app-fiokomdir")= sUserName
      Item.UserProperties.Find("app-datekd")=Cstr(Date())+"-"+Cstr(Time())
      Item.UserProperties.Find("app-optbtnbuff")="OptionButton5"
      Item.UserProperties.Find("Стадия")="Финансовый директор"
      Item.UserProperties.Find("app-savedkd")="Saved"      
   ElseIf OptionButton6.Value=True Then
      mails = "name@domain.ru"
      Call SendMail(mails)
      MsgBox("Финансовому директору отправлено оповещение")
      sUserName = Item.UserProperties.Find("app-userName")
      Item.UserProperties.Find("app-fiokomdir")= sUserName
      Item.UserProperties.Find("app-datekd")=Cstr(Date())+"-"+Cstr(Time())
      Item.UserProperties.Find("app-optbtnbuff")="OptionButton6"
      Item.UserProperties.Find("Стадия")="Финансовый директор"
      Item.UserProperties.Find("app-savedkd")="Saved"      
   End If 
End If
On Error Resume Next
Item.Save

End Sub

Sub CommandButton6_Click() 'Сброс опций Комерческий
  Item.UserProperties.Find("app-optbtnbuff")=""
  Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton4").value = False
  Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton5").value = False
  Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton6").value = False
  Item.UserProperties.Find("app-savedkd")="New"
End Sub

Sub CommandButtonGoToRop_Click()

sFrom = "name@domain.ru"
sText = InputBox("Введите комментарий")
sUserRop = Item.UserProperties.Find("app-rukprodaj")
sLdp = "LDAP://CN="+sUserRop+",ou=staff,ou=moscow,dc=domain,dc=ru" 
On Error Resume Next 
Set objRop = GetObject(sLdp)
mails = objRop.mail
Call SendMail4(mails, sText, sFrom)
Item.UserProperties.Find("Стадия")="РОП"
On Error Resume Next
Item.Save

End Sub

'-----------------------------------------------------------
Sub ButtonPometka9_Click() '**********Снятие пометки удаления*********

    Item.UserProperties.Find("Статус заявки")="Создана"
    On Error Resume Next
    Item.Save
    
End Sub
'-----------------------------------------------------------

Sub CommandButton4_Click() '*********Сохранение************
On Error Resume Next
Item.Save

If Item.UserProperties.Find("app-saveitem")="New" Then
  
   Item.UserProperties.Find("app-saveitem")="Saved"
   
End If 
End Sub
'----------------------------Кредитный контролер-----------------------------------------------
Sub ButtonSaveKK_Click()

    If Item.UserProperties.Find("app-CheckBox4")=True And Item.UserProperties.Find("app-savedKK")="New" Then 
       sUserRop = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("ComboBox3").Value
       sLdp = "LDAP://CN="+sUserRop+",ou=staff,ou=moscow,dc=domain,dc=ru" 
       On Error Resume Next 
       Set objRop = GetObject(sLdp)
       mails = objRop.mail
       Call SendMail(mails)   
      'On Error Resume Next
      'Item.Save
       MsgBox("Заявка отправлена РОП") 
       Item.UserProperties.Find("Стадия")="РОП"
       Item.UserProperties.Find("app-kreditkontrfio")= Item.UserProperties.Find("app-userName")
       Item.UserProperties.Find("app-datekreditkontr")=Cstr(Date())+"-"+Cstr(Time())
       Item.UserProperties.Find("app-savedKK")="Saved"
       On Error Resume Next
       Item.Save
    Else    
       On Error Resume Next
       Item.Save
    End If 
    
End Sub 

Sub CommandButtonKKClose_Click()
    Item.UserProperties.Find("Стадия")="ОПИО"
    Item.UserProperties.Find("Статус заявки")="Завершена"
    On Error Resume Next
    Item.Save
End Sub

Sub CommandButtonResend_Click()
   If Item.UserProperties.Find("app-resend")= "Cупервайзер" Then
       sText = InputBox("Введите комментарий")
       mails = Item.UserProperties.Find("app-mailbuff")
       If mails <> "" Then 'Empty
          Call SendMail5(mails, sText)
          MsgBox("Супервайзеру отправлено оповещение")
          Item.UserProperties.Find("Стадия")="Супервайзер"
          On Error Resume Next
          Item.Save
        Else MsgBox("Поле Супервайзер не заполнено или заполнено неправильно")
        End If 
   ElseIf Item.UserProperties.Find("app-resend")= "РОП" Then
       sText = InputBox("Введите комментарий")
       sUserRop = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("ComboBox3").Value
       sLdp = "LDAP://CN="+sUserRop+",ou=staff,ou=moscow,dc=domain,dc=ru" 
       On Error Resume Next 
       Set objRop = GetObject(sLdp)
       mails = objRop.mail
       Call SendMail5(mails, sText)
       MsgBox("РОПу отправлено оповещение")
       Item.UserProperties.Find("Стадия")="РОП"
       On Error Resume Next
       Item.Save
    ElseIf Item.UserProperties.Find("app-resend")= "Коммерческий директор" Then
       mails = "name@domain.ru" 
       sText = InputBox("Введите комментарий")
       Call SendMail5(mails, sText)
       MsgBox("Коммерческому директору отправлено оповещение")
       Item.UserProperties.Find("Стадия")="Комерческий директор"
       Item.UserProperties.Find("app-savedkd")="New"
       Item.UserProperties.Find("app-optbtnbuff")=""
       On Error Resume Next
       Item.Save      
    ElseIf Item.UserProperties.Find("app-resend")= "Торговый представитель" Then
       sText = InputBox("Введите комментарий")
       sUserTorg = Item.UserProperties.Find("app-torgpredstavfio")
       sLdp = "LDAP://CN="+sUserTorg+",ou=staff,ou=moscow,dc=domain,dc=ru" 
       On Error Resume Next 
       Set objRop = GetObject(sLdp)
       mails = objRop.mail
       Call SendMail5(mails, sText)
       MsgBox("Торговому отправлено оповещение")
       Item.UserProperties.Find("Стадия")="Торговый представитель"
       On Error Resume Next
    ElseIf Item.UserProperties.Find("app-resend")= "Финансовый директор" Then
       mails = "name@domain.ru" 
       sText = InputBox("Введите комментарий")
       Call SendMail5(mails, sText)
       MsgBox("Финансовому директору отправлено оповещение")
       Item.UserProperties.Find("Стадия")="Финансовый директор"
       On Error Resume Next
       Item.Save      
    End If 
       
End Sub
'---------------------------Фин дир------------------------------------------------------------
Sub ButtonSaveFin_Click()

If Item.UserProperties.Find("app-CheckBox7")=True And Item.UserProperties.Find("app-saved6")="New" Then

    If Item.UserProperties.Find("app-company")="ComName1" Or Item.UserProperties.Find("app-company")="ComName1" Then 
      mails = "name@domain.ru"  'К тому кто списывает
    ElseIf Item.UserProperties.Find("app-company")="ComName2" Or Item.UserProperties.Find("app-company")="ComName2" Then  
      mails = "name@domain.ru"
    ElseIf Item.UserProperties.Find("app-company")="ComName3" Or Item.UserProperties.Find("app-company")="ComNmae3" Or Item.UserProperties.Find("app-company")="АТС" Or Item.UserProperties.Find("app-company")="АТТ" Then 
      mails = "name@domain.ru"    
    ElseIf Item.UserProperties.Find("app-company")="ComName4" Or Item.UserProperties.Find("app-company")="ComName4" Then 
      mails = "name@domain.ru" 
End If
  
Call SendMail(mails)
Item.UserProperties.Find("app-findirfio")= Item.UserProperties.Find("app-userName")
Item.UserProperties.Find("app-findirdate")=Cstr(Date())+"-"+Cstr(Time())
Item.UserProperties.Find("Статус заявки")="Утверждена"
Item.UserProperties.Find("Стадия")="ОПИО"
Item.UserProperties.Find("app-saved6")="Saved"

End If
    On Error Resume Next
    Item.Save  
End Sub

Sub CommandButton5_Click()  'Финдир посылает Комерческому   
   mails = "name@domain.ru" 
   sText = InputBox("Введите комментарий")
   Call SendMail2(mails, sText)
   MsgBox("Коммерческому директору отправлено оповещение")
   Item.UserProperties.Find("Стадия")="Комерческий директор"
   Item.UserProperties.Find("app-savedkd")="New"
   Item.UserProperties.Find("app-optbtnbuff")=""
   On Error Resume Next
   Item.Save  
End Sub

'--------------------------Торговый представитель---------------------------------------------
Sub ButtonSaveTp_Click()
   If Item.UserProperties.Find("app-saved1")="New"  Then
      Item.UserProperties.Find("app-CheckBox1")=True
      Item.UserProperties.Find("app-torgpredstavfio")= Item.UserProperties.Find("app-userName")
      Item.UserProperties.Find("app-datetorgpredst")=Cstr(Date())+"-"+Cstr(Time())
      Item.UserProperties.Find("Стадия")="Супервайзер"
      sUserSuper = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("ComboBoxSuper").Value
      sLdp = "LDAP://CN="+sUserSuper+",ou=staff,ou=moscow,dc=domain,dc=ru" 
      On Error Resume Next 
      Set objSuper = GetObject(sLdp)
      sMailSuper = objSuper.mail
      Call SendMail(sMailSuper)
      MsgBox("Супервайзеру отправлено оповещение")
      Item.UserProperties.Find("app-saved1")="Saved"
   End If
   If Item.UserProperties.Find("Стадия")="Торговый представитель" Then  
      sUserSuper = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("ComboBoxSuper").Value
      sLdp = "LDAP://CN="+sUserSuper+",ou=staff,ou=moscow,dc=domain,dc=ru" 
      On Error Resume Next 
      Set objSuper = GetObject(sLdp)
      sMailSuper = objSuper.mail
      Call SendMail(sMailSuper)
      MsgBox("Супервайзеру отправлено оповещение")
      Item.UserProperties.Find("app-datetorgpredst")=Cstr(Date())+"-"+Cstr(Time())
      Item.UserProperties.Find("Стадия")="Супервайзер"      
   End If 
   On Error Resume Next
   Item.Save
End Sub
'--------------------------Супервайзер--------------------------------------------------------
Sub ButtonSaveSuper_Click()
    If Item.UserProperties.Find("app-saved2")="New" Then
       If Item.UserProperties.Find("app-rop") = "Rop name" Then        
          mails = "name@domain.ru"
          Call SendMail(mails)    
       ElseIf Item.UserProperties.Find("app-rop") = "Rop name" Then
          mails = "name@domain.ru"
          Call SendMail(mails)  
       End If 
       uName = Item.UserProperties.Find("app-userName")
       Item.UserProperties.Find("app-supervfio")= uName
       Item.UserProperties.Find("app-datesupervis")=Cstr(Date())+"-"+Cstr(Time())
       Item.UserProperties.Find("Стадия")="Кредитный контролер"
       Item.UserProperties.Find("app-saved2")="Saved"
     End If
     If Item.UserProperties.Find("Стадия")="Супервайзер" Then  
       If Item.UserProperties.Find("app-rop") = "Rop name" Then        
          mails = "name@domain.ru"
          Call SendMail(mails)    
       ElseIf Item.UserProperties.Find("app-rop") = "Rop name" Then
          mails = "name@domain.ru"
          Call SendMail(mails)  
       End If 
       Item.UserProperties.Find("app-datesupervis")=Cstr(Date())+"-"+Cstr(Time())
       Item.UserProperties.Find("Стадия")="Кредитный контролер"    
     End If 
     On Error Resume Next
     Item.Save
End Sub
'---------------------------РОП--------------------------------------------------------------
Sub ButtonSaveRop_Click()
    If Item.UserProperties.Find("app-saved3")="New" Then
       mails = "name@domain.ru"
       Call SendMail(mails)
       Item.UserProperties.Find("app-CheckBox3")=True
       Item.UserProperties.Find("app-rukprodaj")= Item.UserProperties.Find("app-userName")
       Item.UserProperties.Find("app-daterukprod")=Cstr(Date())+"-"+Cstr(Time())
       Item.UserProperties.Find("Стадия")="Комерческий директор"
       Item.UserProperties.Find("app-saved3")="Saved"
End If
    If Item.UserProperties.Find("Стадия")="РОП" Then
       mails = "name@domain.ru"
       Call SendMail(mails)
       Item.UserProperties.Find("app-toSB")="New" 'Сброс у Д
       Item.UserProperties.Find("app-komdircomment")="" 
       Item.UserProperties.Find("Стадия")="Коммерческий директор"
    End If 
    On Error Resume Next
    Item.Save
End Sub

Sub ButtonGoToSuper_Click()
    sText = InputBox("Введите комментарий")
    sUserSuper = Item.UserProperties.Find("app-supervfio")
    sLdp = "LDAP://CN="+sUserSuper+",ou=staff,ou=moscow,dc=domain,dc=ru" 
    On Error Resume Next 
    Set objRop = GetObject(sLdp)
    mails = objRop.mail
    Call SendMail2(mails, sText)
    Item.UserProperties.Find("Стадия")="Супервайзер"
    On Error Resume Next
    Item.Save
End Sub

Sub CommandButton2_Click()
   sText = "Вы точно хотите пометить на удаление? Заявка будет заблокирована"
   rc = MsgBox(sText,4) 
   If rc = 6 Then 
   Item.UserProperties.Find("Статус заявки")="Удалено"
   On Error Resume Next
   Item.Save  
   End If  
End Sub
'--------------------------Юристы-------------------------------------------------------------
Sub CommandButtonPrin_Click()
   If Item.UserProperties.Find("app-saveddpu") = "New"  Then
       Item.UserProperties.Find("app-datepostupl")=Cstr(Date())+"-"+Cstr(Time())
       Item.UserProperties.Find("Стадия")="Принято в работу Юр отд"
       Item.UserProperties.Find("app-urprin")= Item.UserProperties.Find("app-userName")
   '-------Юр номера----------
   Set ofso1 = CreateObject("Scripting.FileSystemObject") 'Объект файловой системы
   sTargetFolder = "\\your-file-server\systemf$\NumberUr.txt"   'Файл с номерами
   Set oFile=ofso1.GetFile(sTargetFolder)
   num = 1
   nCount = 0
   If ofso1.FileExists(sTargetFolder) Then            
           Set oFile = ofso1.OpenTextFile(sTargetFolder,1,True)        
           Do While oFile.AtEndOfStream <> True
               oFile.SkipLine
                   nCount = nCount + 1
           Loop 
           If  nCount = 0 Then 
                Item.UserProperties.Find("Юр номер") = Cstr(nCount+1)             
           ElseIf nCount <> 0 Then
                 Item.UserProperties.Find("Юр номер") =Cstr(nCount+1)                               
           End If   
   End If

   Set oFile = ofso1.OpenTextFile(sTargetFolder,8,True)
   oFile.WriteLine(Item.UserProperties.Find("Юр номер"))
   '-----------------

       Item.UserProperties.Find("app-saveddpu") = "Saved" 
   End If
   
   On Error Resume Next
   Item.Save 
End Sub

Sub ButtonSaveUrist_Click()
    If Item.UserProperties.Find("app-uristcomment")<>"" And Item.UserProperties.Find("app-datepostupl")<>"" And Item.UserProperties.Find("app-CheckBox6")<>True And Item.UserProperties.Find("app-CheckBox5")<>True Then 
          sUser1 = Item.UserProperties.Find("app-torgpredstavfio")
          sLdp1 = "LDAP://CN="+sUser1+",ou=staff,ou=moscow,dc=domain,dc=ru"
          Set objU1 = GetObject(sLdp1)
          sUserSuper = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("ComboBoxSuper").Value
          sLdp = "LDAP://CN="+sUserSuper+",ou=staff,ou=moscow,dc=domain,dc=ru" 
          On Error Resume Next 
          Set objSuper = GetObject(sLdp)
          sMailSuper = objSuper.mail
          sMail1 = objU1.mail 'Мыло манагера
          sMail2 = sMailSuper 'Мыло супервайзера      
          mails = sMail1+";"+sMail2 
          sText = Item.UserProperties.Find("app-uristcomment")
          Call SendMail3(mails, sText)
          MsgBox("Комментарий отправлен супервайзеру и торговому представителю")
    End If 
    
    On Error Resume Next
    Item.Save   
End Sub

Sub CommandButtonBack_Click()
    
    Item.UserProperties.Find("Стадия")="Торговый представитель"
    On Error Resume Next
    Item.Save  
    
    sUserSuper = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("ComboBoxSuper").Value
    sLdp = "LDAP://CN="+sUserSuper+",ou=staff,ou=moscow,dc=domain,dc=ru" 
    On Error Resume Next 
    Set objSuper = GetObject(sLdp)
    sMailSuper = objSuper.mail
    sUser1 = Item.UserProperties.Find("app-rukprodaj")
    sLdp1 = "LDAP://CN="+sUser1+",ou=staff,ou=moscow,dc=domain,dc=ru"
    'MsgBox(sLdp1)
    Set objU1 = GetObject(sLdp1)
    sMail1 = objU1.mail
    sMail2 = sMailSuper
    'MsgBox(sMail2)
    'MsgBox(sMail1)
     Dim objOL   'As Outlook.Application
     Dim objAppt 'As Outlook.AppointmentItem
     Const olAppointmentItem = 1
     Const olMeeting = 1
     
      Set myItem = Application.CreateItem(olAppointmentItem) 
       myItem.MeetingStatus = olMeeting      
      myItem.RequiredAttendees = sMail1 + ";" + sMail2 'Событие супервайзеру о непринятом док-те      
       myItem.Subject = "Заявка на списание дебиторской задолженности №" + Item.UserProperties.Find("Номер заявки") + " не принята в работу юридическим отделом" 
       myItem.Location = "Автор: " + Item.UserProperties.Find("app-manfio") + " - " + Item.UserProperties.Find("app-uristcomment")
       myItem.Start = Date()
       myItem.ReminderSet = True
       myItem.ReminderMinutesBeforeStart = 90
      myItem.end = DateAdd("d",10,Date)               
       'myItem.Display
       myItem.Send
End Sub

Sub CommandButton7_Click() 'Сброс опций
  Item.UserProperties.Find("app-optbtnbuff1")=""
  Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton1").value = False
  Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton2").value = False
  Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton3").value = False
  Item.UserProperties.Find("app-saved7")="New"
End Sub
'----------------------Кнопка кредитному------------------------------------------------------------
Sub CommandButton8_Click()
    mails = "name@domain.ru;name@domain.ru"
    Call SendMail(mails)    
    'On Error Resume Next
    'Item.Save
    MsgBox("Отправлен запрос кредитному контролеру")
End Sub
'===========================================================================================================================================
'--------------------------Административный директор Н Ю.А.---------------------------
Sub CommandButtonPrin10_Click() 'Принять в работу
   If Item.UserProperties.Find("app-saveddpa") = "New"  Then
       Item.UserProperties.Find("app-datapostuplAD")=Cstr(Date())+"-"+Cstr(Time())
       Item.UserProperties.Find("Стадия")="Принято в работу Адм. дир."
       Item.UserProperties.Find("app-saveddpa") = "Saved" 
   End If
   
   On Error Resume Next
   Item.Save 
End Sub

Sub CommandButton9_Click() 'Адм дир посылает в СБ
   If Item.UserProperties.Find("app-ADtoSB")="New" Then 
      
       MsgBox("Юристам отправлено оповещение")
       Item.UserProperties.Find("app-fioAD")= Item.UserProperties.Find("app-userName")
       Item.UserProperties.Find("app-dateAD1")=Cstr(Date())+"-"+Cstr(Time())
       Item.UserProperties.Find("Стадия")="Юридический отдел"
       Item.UserProperties.Find("app-ADtoSB")="Saved"
            
   End If
   On Error Resume Next
   Item.Save
End Sub


Sub CommandButtonSaveAD_Click() 'Кнопка "Сохранить\Списать"

If Item.UserProperties.Find("app-savedAD")="New" Then
   Set OptionButton7 = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton7")
   Set OptionButton8 = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton8")
   
   If OptionButton7.Value=True Then
      mails = "name@domain.ru"
      Call SendMail(mails)
      MsgBox("Финансовому директору отправлено оповещение")
      Item.UserProperties.Find("app-fioAD")= Item.UserProperties.Find("app-userName")
      Item.UserProperties.Find("app-dateAD1")=Cstr(Date())+"-"+Cstr(Time())
      Item.UserProperties.Find("app-optbtnbuff5")="OptionButton7"
      Item.UserProperties.Find("Стадия")="Финансовый директор"
      Item.UserProperties.Find("app-savedAD")="Saved"      
   ElseIf OptionButton8.Value=True Then
      mails = "name@domain.ru"
      Call SendMail(mails)
      MsgBox("Финансовому директору отправлено оповещение")
      Item.UserProperties.Find("app-fioAD")= Item.UserProperties.Find("app-userName")
      Item.UserProperties.Find("app-dateAD1")=Cstr(Date())+"-"+Cstr(Time())
      Item.UserProperties.Find("app-optbtnbuff5")="OptionButton8"
      Item.UserProperties.Find("Стадия")="Финансовый директор"
      Item.UserProperties.Find("app-savedAD")="Saved"        
   End If 
   
End If
On Error Resume Next
Item.Save

End Sub

Sub CommandButtonBack11_Click() 'Кнопка "Вернуть торговому"
    
    Item.UserProperties.Find("Стадия")="Торговый представитель"
    MsgBox("Заявка отправлена торговому представителю")
    On Error Resume Next
    Item.Save  
    sUserSuper = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("ComboBoxSuper").Value
    sLdp = "LDAP://CN="+sUserSuper+",ou=staff,ou=moscow,dc=domain,dc=ru" 
    On Error Resume Next 
    Set objSuper = GetObject(sLdp)
    sMailSuper = objSuper.mail
    sUser1 = Item.UserProperties.Find("app-rukprodaj")
    sLdp1 = "LDAP://CN="+sUser1+",ou=staff,ou=moscow,dc=domain,dc=ru"
    'MsgBox(sLdp1)
    Set objU1 = GetObject(sLdp1)
    sMail1 = objU1.mail
    sMail2 = sMailSuper
    'MsgBox(sMail2)
    'MsgBox(sMail1)
     Dim objOL   'As Outlook.Application
     Dim objAppt 'As Outlook.AppointmentItem
     Const olAppointmentItem = 1
     Const olMeeting = 1
     
      Set myItem = Application.CreateItem(olAppointmentItem) 
       myItem.MeetingStatus = olMeeting      
      myItem.RequiredAttendees = sMail1 + ";" + sMail2 'Событие супервайзеру о непринятом док-те      
       myItem.Subject = "Заявка на списание дебиторской задолженности №" + Item.UserProperties.Find("Номер заявки") + " не принята в работу Административным директором" 
       myItem.Location = "Автор: " + Item.UserProperties.Find("app-manfio") + " - " + Item.UserProperties.Find("app-uristcomment")
       myItem.Start = Date()
       myItem.ReminderSet = True
       myItem.ReminderMinutesBeforeStart = 90
      myItem.end = DateAdd("d",10,Date)               
       'myItem.Display
       myItem.Send
End Sub

Sub CommandButtonDell9_Click() 'Кнопка Сброс опций АД
  Item.UserProperties.Find("app-optbtnbuff5")=""
  Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton7").value = False
  Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton8").value = False
  Item.UserProperties.Find("app-savedAD")="New"
End Sub

'===========================================================================================================================================
'----------------------Печать-----------------------------------------------------------------------
Sub cmdPrint_Click()
   Set oWordApp = CreateObject("Word.Application")
   If oWordApp Is Nothing Then
      MsgBox "Couldn't start Word."
   Else
      Dim oWordApp, nCol, n
      Dim oWordDoc
      Dim bolPrintBackground
	Dim strMyField1,strMyField2,strMyField3,strMyField4,strMyField5,strMyField6a,strMyField6b,strMyField7,strMyField8
      Dim strMyField9,strMyField10,strMyField11,strMyField12,strMyField13,strMyField14,strMyField15,strMyField16

      ' Open a new document
      Set oDoc = oWordApp.Documents.Add("\\domain.ru\NETLOGON\MyFormDeb10.dot")
   	
    If Item.UserProperties.Find("Юр номер")<>"" And Item.UserProperties.Find("app-buffgroup") = "urotdel" Then
    strMyField1 = Item.UserProperties.Find("Юр номер")
        oDoc.FormFields("ТекстовоеПоле1").Result = strMyField1
    Else 
    strMyField1 = Item.UserProperties.Find("Номер заявки")
       	oDoc.FormFields("ТекстовоеПоле1").Result = strMyField1
    End If 
    strMyField2 = Item.UserProperties.Find("app-company")
       	oDoc.FormFields("ТекстовоеПоле2").Result = strMyField2

	strMyField3 = Item.UserProperties.Find("app-manfio")
       	oDoc.FormFields("ТекстовоеПоле3").Result = strMyField3

	strMyField4 = Item.UserProperties.Find("app-nazvclient")
       	oDoc.FormFields("ТекстовоеПоле4").Result = strMyField4

	strMyField5 = Item.UserProperties.Find("app-kodclient")
	    oDoc.FormFields("ТекстовоеПоле5").Result = strMyField5
	
	strMyField6 = Item.UserProperties.Find("app-faktnazclient")
	    oDoc.FormFields("ТекстовоеПоле6").Result = strMyField6
	
	strMyField7 = Item.UserProperties.Find("app-contclient")
       	oDoc.FormFields("ТекстовоеПоле7").Result = strMyField7

	strMyField8 = Item.UserProperties.Find("app-summvsego")
       	oDoc.FormFields("ТекстовоеПоле8").Result = strMyField8

	strMyField9 = Item.UserProperties.Find("app-summbuh")
       	oDoc.FormFields("ТекстовоеПоле9").Result = strMyField9

	strMyField10 = Item.UserProperties.Find("app-summupr")
       	oDoc.FormFields("ТекстовоеПоле10").Result = strMyField10

	strMyField11 = Item.UserProperties.Find("app-datezadol")
       	oDoc.FormFields("ТекстовоеПоле11").Result = strMyField11

	strMyField12 = Item.UserProperties.Find("app-summmntar")
       	oDoc.FormFields("ТекстовоеПоле12").Result = strMyField12

	strMyField13 = Item.UserProperties.Find("app-summtorgobor")
       	oDoc.FormFields("ТекстовоеПоле13").Result = strMyField13

'-------------------------Торговый печать------------------------------------------------
    strMyField14 = Item.UserProperties.Find("app-commenttorg")
    oDoc.FormFields("ТекстовоеПоле14").Range.InsertAfter strMyField14
    
	strMyField15 = Item.UserProperties.Find("app-torgpredstavfio")
       	oDoc.FormFields("ТекстовоеПоле15").Result = strMyField15

	strMyField16 = Item.UserProperties.Find("app-datetorgpredst")
       	oDoc.FormFields("ТекстовоеПоле16").Result = strMyField16

	strMyFieldFl1 = Item.UserProperties.Find("app-CheckBox1")
	If strMyFieldFl1 = True Then
       	oDoc.FormFields("Флажок1").CheckBox.Value = strMyFieldFl1
    End If  
'-------------------------Супервайзер----------------------------------------------------
    strMyField17 = Item.UserProperties.Find("app-commentsuperv") 
    oDoc.FormFields("ТекстовоеПоле17").Range.InsertAfter strMyField17
        
	strMyField18 = Item.UserProperties.Find("app-supervfio")
       	oDoc.FormFields("ТекстовоеПоле18").Result = strMyField18

	strMyField19 = Item.UserProperties.Find("app-datesupervis")
       	oDoc.FormFields("ТекстовоеПоле19").Result = strMyField19

	strMyFieldFl2 = Item.UserProperties.Find("app-CheckBox2")
	If strMyFieldFl2 = True Then
       	oDoc.FormFields("Флажок2").CheckBox.Value = strMyFieldFl2
    End If  
'-----------------------------РОП--------------------------------------------------------
    strMyField20 = Item.UserProperties.Find("app-commentrukprod")
    If Len(strMyField20) > 250 Then
        nCol = Len(strMyField20) 
        oDoc.FormFields("ТекстовоеПоле20").Result = Left(strMyField20, 250) 
        oDoc.FormFields("ТекстовоеПоле45").Result = Right(strMyField20, nCol - 250)  	
    Else   	
       	oDoc.FormFields("ТекстовоеПоле20").Result = strMyField20
    End If
	strMyField21 = Item.UserProperties.Find("app-rukprodaj")
       	oDoc.FormFields("ТекстовоеПоле21").Result = strMyField21

	strMyField22 = Item.UserProperties.Find("app-daterukprod")
       	oDoc.FormFields("ТекстовоеПоле22").Result = strMyField22

	strMyFieldFl3 = Item.UserProperties.Find("app-CheckBox3")
	If strMyFieldFl3 = True Then
       	oDoc.FormFields("Флажок3").CheckBox.Value = strMyFieldFl3
    End If  
'-------------------------------КомДир---------------------------------------------------  
    strMyField23 = Item.UserProperties.Find("app-komdircomment")
    If Len(strMyField23) > 250 Then
        nCol = Len(strMyField23) 
        oDoc.FormFields("ТекстовоеПоле23").Result = Left(strMyField23, 250) 
        oDoc.FormFields("ТекстовоеПоле46").Result = Right(strMyField23, nCol - 250)  	
    Else   	       	     	
       	oDoc.FormFields("ТекстовоеПоле23").Result = strMyField23
    End If
	strMyField24 = Item.UserProperties.Find("app-fiokomdir")
       	oDoc.FormFields("ТекстовоеПоле24").Result = strMyField24

	strMyField25 = Item.UserProperties.Find("app-datekd")
       	oDoc.FormFields("ТекстовоеПоле25").Result = strMyField25
       	   	
   If Item.UserProperties.Find("app-optbtnbuff")="OptionButton4" Then 
       	oDoc.FormFields("ТекстовоеПоле26").Result = "ЗП менеджера" 
   ElseIf Item.UserProperties.Find("app-optbtnbuff")="OptionButton5" Then 
        oDoc.FormFields("ТекстовоеПоле26").Result = "Фонда"
   ElseIf Item.UserProperties.Find("app-optbtnbuff")="OptionButton6" Then
       oDoc.FormFields("ТекстовоеПоле26").Result = "Компании"
   End If 	      	
'------------------------------Кредитный---27-32------------------------------------------
    
    strMyField27 = Item.UserProperties.Find("app-otgruzpvpz")
       	oDoc.FormFields("ТекстовоеПоле27").Result = strMyField27

	strMyField28 = Item.UserProperties.Find("app-otgruzsl")
       	oDoc.FormFields("ТекстовоеПоле28").Result = strMyField28

	strMyField29 = Item.UserProperties.Find("app-istorrabskl")
       	oDoc.FormFields("ТекстовоеПоле29").Result = strMyField29
    
    strMyField30 = Item.UserProperties.Find("app-commentkreditkontr")
    If Len(strMyField30) > 250 Then
        nCol = Len(strMyField30) 
        oDoc.FormFields("ТекстовоеПоле30").Result = Left(strMyField30, 250) 
        oDoc.FormFields("ТекстовоеПоле47").Result = Right(strMyField30, nCol - 250)  	
    Else               
       	oDoc.FormFields("ТекстовоеПоле30").Result = strMyField30
    End If
	strMyField31 = Item.UserProperties.Find("app-kreditkontrfio")
       	oDoc.FormFields("ТекстовоеПоле31").Result = strMyField31

	strMyField32 = Item.UserProperties.Find("app-datekreditkontr")
       	oDoc.FormFields("ТекстовоеПоле32").Result = strMyField32	

	strMyFieldFl4 = Item.UserProperties.Find("app-CheckBox4")
	If strMyFieldFl4 = True Then
       	oDoc.FormFields("Флажок4").CheckBox.Value = strMyFieldFl4
    End If  
'---------------------------Юристы---33-39---------------------------------------------
    strMyField33 = Item.UserProperties.Find("app-datepostupl")
       	oDoc.FormFields("ТекстовоеПоле33").Result = strMyField33

	strMyField34 = Item.UserProperties.Find("app-uristcomment")
    If Len(strMyField34) > 250 Then
        nCol = Len(strMyField34) 
        oDoc.FormFields("ТекстовоеПоле34").Result = Left(strMyField34, 250) 
        oDoc.FormFields("ТекстовоеПоле48").Result = Right(strMyField34, nCol - 250)  	
    Else 	       	
       oDoc.FormFields("ТекстовоеПоле34").Result = strMyField34
    End If
	strMyField35 = Item.UserProperties.Find("app-sotrurotdel")
       	oDoc.FormFields("ТекстовоеПоле35").Result = strMyField35
    
    strMyField36 = Item.UserProperties.Find("app-dateurist1")
       	oDoc.FormFields("ТекстовоеПоле36").Result = strMyField36

	strMyField37 = Item.UserProperties.Find("app-rukovurotdel")
       	oDoc.FormFields("ТекстовоеПоле37").Result = strMyField37

	strMyField38 = Item.UserProperties.Find("app-dateurist2")
       	oDoc.FormFields("ТекстовоеПоле38").Result = strMyField38	
    
    If Item.UserProperties.Find("app-optbtnbuff1")="OptionButton1" Then
       oDoc.FormFields("ТекстовоеПоле39").Result = "Списать"
    ElseIf Item.UserProperties.Find("app-optbtnbuff1")="OptionButton2" Then
       oDoc.FormFields("ТекстовоеПоле39").Result = "Оплачено 100%"
    ElseIf Item.UserProperties.Find("app-optbtnbuff1")="OptionButton3" Then
       oDoc.FormFields("ТекстовоеПоле39").Result = "Возврат из СБ"
    End If 
       		
    
	strMyFieldFl5 = Item.UserProperties.Find("app-CheckBox5")
	If strMyFieldFl5 = True Then
       	oDoc.FormFields("Флажок5").CheckBox.Value = strMyFieldFl5
    End If  
    strMyFieldFl6 = Item.UserProperties.Find("app-CheckBox6")
	If strMyFieldFl6 = True Then
       	oDoc.FormFields("Флажок6").CheckBox.Value = strMyFieldFl6
    End If  
'----------------------------Финдир---40-42--------------------------------------------------------
    strMyField40 = Item.UserProperties.Find("app-findirreshen")
    If Len(strMyField40) > 250 Then
        nCol = Len(strMyField40) 
        oDoc.FormFields("ТекстовоеПоле40").Result = Left(strMyField40, 250) 
        oDoc.FormFields("ТекстовоеПоле49").Result = Right(strMyField40, nCol - 250)  	
    Else           
       	oDoc.FormFields("ТекстовоеПоле40").Result = strMyField40
    End If
	strMyField41 = Item.UserProperties.Find("app-findirfio")
       	oDoc.FormFields("ТекстовоеПоле41").Result = strMyField41

	strMyField42 = Item.UserProperties.Find("app-findirdate")
       	oDoc.FormFields("ТекстовоеПоле42").Result = strMyField42

	strMyFieldFl7 = Item.UserProperties.Find("app-CheckBox7")
	If strMyFieldFl7 = True Then
       	oDoc.FormFields("Флажок7").CheckBox.Value = strMyFieldFl7
    End If  
'----------------------------------------------------------------------------------------

      ' Get the current Word setting for background printing
      bolPrintBackground = oWordApp.Options.PrintBackground

      ' Turn background printing off
      oWordApp.Options.PrintBackground = False

      ' Print the Word document
      oDoc.PrintOut

      ' Restore previous setting
      oWordApp.Options.PrintBackground = bolPrintBackground

      ' Close and don't save changes to the document
      Const wdDoNotSaveChanges = 0
      oDoc.Close wdDoNotSaveChanges

      ' Close the Word instance
      oWordApp.Quit

      ' Clean up
      Set oDoc = Nothing
      Set oWordApp = Nothing
   End If
End Sub
    

'*************************************************************************************************************************************************

'*************************************************************************************************************************************************
Function Item_Write()

'---------------------------Служебная записка-------------------------------------------------------------
If Item.UserProperties.Find("app-company") = "" Then
    MsgBox("Поле должно быть заполнено!")
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("ComboBox1").BackColor = RGB(255,0,0)
    Item_Write = False
ElseIf  Item.UserProperties.Find("app-manfio") = "" Then
    MsgBox("Поле должно быть заполнено!")
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox1").BackColor = RGB(255,0,0)
    Item_Write = False
ElseIf Item.UserProperties.Find("app-nazvclient") = "" Then
    MsgBox("Поле должно быть заполнено!")
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox2").BackColor = RGB(255,0,0)
    Item_Write = False 
ElseIf Item.UserProperties.Find("app-kodclient") = "" Then
    MsgBox("Поле должно быть заполнено!")
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox3").BackColor = RGB(255,0,0)
    Item_Write = False
ElseIf Item.UserProperties.Find("app-faktnazclient") = "" Then 
    MsgBox("Поле должно быть заполнено!")
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox4").BackColor = RGB(255,0,0)
    Item_Write = False
ElseIf Item.UserProperties.Find("app-contclient") = ""   Then
    MsgBox("Поле должно быть заполнено!")
    Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("TextBox5").BackColor = RGB(255,0,0)
    Item_Write = False   
End If


'--------------------Письмо о создании------------------------------------------------------------------------------

If Item.UserProperties.Find("app-saved")="New" Then


'-----------------
Set ofso1 = CreateObject("Scripting.FileSystemObject") 'Объект файловой системы
sTargetFolder = "\\your-file-serversystemf$\Number.txt"   'Файл с номерами
Set oFile=ofso1.GetFile(sTargetFolder)

num = 1
nCount = 0
If ofso1.FileExists(sTargetFolder) Then            
           Set oFile = ofso1.OpenTextFile(sTargetFolder,1,True)        
           Do While oFile.AtEndOfStream <> True
               oFile.SkipLine
                   nCount = nCount + 1
           Loop 
           If  nCount = 0 Then 
                Item.UserProperties.Find("Номер заявки") = Cstr(nCount+1)             
           ElseIf nCount <> 0 Then
                 Item.UserProperties.Find("Номер заявки") =Cstr(nCount+1)                               
           End If   
End If

Set oFile = ofso1.OpenTextFile(sTargetFolder,8,True)
oFile.WriteLine(Item.UserProperties.Find("Номер заявки"))
'-----------------

Item.UserProperties.Find("app-mailbuff") = sMail
Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame2").Visible = True
Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("Frame3").Visible = True
'-----------------
Item.UserProperties.Find("Стадия")="Супервайзер"
'-----------------
Item.UserProperties.Find("app-saved")="Saved"

 
End If

'------------------------------Письма в зависимости от прохождения-------------------------------------------------------------

'-------------------Комментарии кредитного контролера-------------------------------


'-----------------------От юристов к финдиру------------------------------


Set OptionButton1 = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton1")
Set OptionButton2 = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton2")
Set OptionButton3 = Item.GetInspector.ModifiedFormPages("Списание оборудования").Controls("OptionButton3")

If  Item.UserProperties.Find("app-CheckBox5")=True And Item.UserProperties.Find("app-saved5")="New" Then     
      Item.UserProperties.Find("app-sotrurotdel")= Item.UserProperties.Find("app-userName")
      Item.UserProperties.Find("app-dateurist1")=Cstr(Date())+"-"+Cstr(Time())
      Item.UserProperties.Find("app-saved5")="Saved"      
End If 

If  Item.UserProperties.Find("app-CheckBox6")=True And Item.UserProperties.Find("app-saved7")="New" Then   
   If OptionButton1.Value=True Then
      mails = "name@domain.ru"
      Call SendMail(mails)
      Item.UserProperties.Find("app-rukovurotdel")= Item.UserProperties.Find("app-userName")
      Item.UserProperties.Find("app-dateurist2")=Cstr(Date())+"-"+Cstr(Time())
      Item.UserProperties.Find("app-optbtnbuff1")="OptionButton1"
      Item.UserProperties.Find("Стадия")="Финансовый директор"
      Item.UserProperties.Find("app-saved7")="Saved"      
   ElseIf OptionButton2.Value=True Then
      
      Item.UserProperties.Find("app-rukovurotdel")= Item.UserProperties.Find("app-userName")
      Item.UserProperties.Find("app-dateurist2")=Cstr(Date())+"-"+Cstr(Time())
      Item.UserProperties.Find("app-optbtnbuff1")="OptionButton2"
      Item.UserProperties.Find("Статус заявки")="Завершена"
      Item.UserProperties.Find("app-saved7")="Saved"      
   ElseIf OptionButton3.Value=True Then
      
      Item.UserProperties.Find("app-rukovurotdel")= Item.UserProperties.Find("app-userName")
      Item.UserProperties.Find("app-dateurist2")=Cstr(Date())+"-"+Cstr(Time())
      Item.UserProperties.Find("app-optbtnbuff1")="OptionButton3"
      Item.UserProperties.Find("Статус заявки")="Завершена"
      Item.UserProperties.Find("app-saved7")="Saved"      
   End If 

End If

'--------------------ОПИО----------------------------

If Item.UserProperties.Find("app-CheckBox9")=True And Item.UserProperties.Find("app-savedclose")="New" Then
Item.UserProperties.Find("Статус заявки")="Завершена"
Item.UserProperties.Find("app-savedclose")="Saved" 
End If

End Function