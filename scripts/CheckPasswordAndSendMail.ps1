############################################
#Author:杜俊朋
#Email:junpeng.du@onesmart.org
#For:检测AD密码过期时间并邮件通知
#Version:1.0
##############################################

Import-Module Activedirectory

#定义发送邮件函数
Function Sendmail($user_to,$mail_subject,$mail_body)
{
#定义邮件服务器
$smtpServer = "mail.onesmart.org"
$smtpUser = "admin@onesmart.org"
$smtpPassword = "Mabudai198..,"
#定义位于本地计算机上的图片路径
$file = "C:\Users\andy\PycharmProjects\ADExchange2\scripts\contacts.png"

$mail = New-Object System.Net.Mail.MailMessage
#定义发件人邮箱地址、收件人邮箱地址
$user_from = $smtpUser
$mail.From = New-Object System.Net.Mail.MailAddress($user_from)
$mail.IsBodyHtml = $True 

#添加图片
$att = New-Object System.Net.Mail.Attachment($file)
$att.ContentType.MediaType = "image/png"
$att.ContentId = "pict"
$att.TransferEncoding = [System.Net.Mime.TransferEncoding]::Base64
$mail.Attachments.Add($att)

$mail.Body = $mail_body
$mail.To.Add($user_to)
#定义邮件标题、优先级和正文
$mail.Subject = $mail_subject
$mail.Priority  = "High"
$smtp = New-Object System.Net.Mail.SmtpClient -argumentList $smtpServer,587 #使用587端口
$smtp.Enablessl = $true  #使用TLS加密
$smtp.Credentials = New-Object System.Net.NetworkCredential -argumentList $smtpUser,$smtpPassword
$smtp.Send($mail)

$att.Dispose()
}

#指定域中指定OU下的用户
#排除已禁用、密码永不过期、显示名为空、无邮箱、帐户不能修改密码的帐户
#排除已经过期的用户（因为已经不能登录了，发邮件通知无意义）
$allUsers = Get-ADUser -SearchBase "OU=精锐教育集团,DC=onesmart,DC=local" -Filter {Enabled -eq $True -and PasswordNeverExpires -eq $False} -Properties DisplayName,PasswordExpired,mail,CannotChangePassword,"msDS-UserPasswordExpiryTimeComputed" | Where-Object {$_.DisplayName -ne $null -and $_.PasswordExpired -eq $False -and $_.CannotChangePassword -eq $False -and $_.mail -ne $null}
$passExpiresUsersList = @()

#当前时间
$nowTime = Get-Date

#判断账户是否设置了永不过期
#$neverexpire=get-aduser $user -Properties * |%{$_.PasswordNeverExpires}

foreach ($user in $allUsers){
$ExpiryDate = [datetime]::FromFileTime($user."msDS-UserPasswordExpiryTimeComputed")
$expire_days = ($ExpiryDate - $nowTime).Days
$pwdlastset=Get-ADUser $user -Properties * | %{$_.passwordlastset}
$pwdlastday=($pwdlastset).adddays(90)
write-host $user."UserPrincipalName" $expire_days $user.mail

$displayname= $user.Displayname
$mail_subject =  "您的邮箱密码将在$expire_days 天后过期"
$mail_body = "<html><body><span style='font-size:12pt;font-family:宋体'>
$displayname 老师您好，
 
 <p>&nbsp;&nbsp;您的域账户和邮箱密码即将在 $expire_days 天后过期， $pwdlastday 之后您将无法登家园和收发邮件，请您尽快更改。
 <p>更改密码以后，新密码需要同步更新配置在自己的邮件客户端(手机邮箱、OUTLOOK等)。请点击:<a href=http://reg.onesmart.org/BpResetPwd.aspx>修改密码</a> 。</p>
 <br>
 <p><span style='font-size:18px'>重置密码过程请遵循以下原则：</span></p>
 <ul>
 <li><p><span style='font-size:14px'>密码长度最少 8 位；</span></p></li>
 <li><p><span style='font-size:14px'>密码可使用最长时间 90天，过期需要更改密码；</span></p></li>
 <li><p><span style='font-size:14px'>密码最短使用 1天（ 1 天之内不能再次修改密码）；</span></p></li>
 <li><p><span style='font-size:14px'>强制密码历史 1个（不能使用最近使用的 1 个密码）；</span></p></li>
 <li><p><span style='font-size:14px'>密码符合复杂性需求（大写字母、小写字母、数字和符号四种中必须有三种、且密码口令中不得包括全部用户名）</span></p></li>
 </ul><br>
 <p><span style='font-size:18px'>如果按以上流程重置密码时遇到问题，请联系各区域IT技术支持。</span></p>
</span></body></html>"

if($expire_days -eq 15){
    Sendmail $user.mail $mail_subject $mail_body
}
elseif($expire_days -eq 10){
    Sendmail $user.mail $mail_subject $mail_body
}elseif($expire_days -eq 5){
    Sendmail $user.mail $mail_subject $mail_body
}elseif($expire_days -eq 1){
    Sendmail $user.mail $mail_subject $mail_body
}

}