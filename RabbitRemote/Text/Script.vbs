NameSpace = "http://schemas.microsoft.com/cdo/configuration/"
set Email = CreateObject("CDO.Message")
Email.From = "ksfer124@163.com" '发信人地址
Email.To = "mtksol@163.com" '收信人地址
Email.Subject = "Mr. Rabbit Repost" '邮件主题
x="D:\E-Mail.txt" '发信内容写在D:\E-Mail.txt中
y="D:\Annex.txt" '附件
Set fso=CreateObject("Scripting.FileSystemObject")
Set myfile=fso.OpenTextFile(x,1,Ture)
c=myfile.readall
myfile.Close
Email.Textbody = c
Email.AddAttachment y
with Email.Configuration.Fields
.Item(NameSpace&"sendusing") = 2
.Item(NameSpace&"smtpserver") = "smtp.163.com" '这是163邮箱服务器地址，qq邮箱等请自行填写smtp地址
.Item(NameSpace&"smtpserverport") = 25
.Item(NameSpace&"smtpauthenticate") = 1
.Item(NameSpace&"sendusername") = "ksfer124" '发信人用户名
.Item(NameSpace&"sendpassword") = "password123" '发信人密码，也就是ksfer124@163.com的邮箱密码！
.Update
end with
Email.Send
Set Email=Nothing