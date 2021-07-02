Import-Module PowerShellAccessControl
Import-Module activedirectory
set-location ad:

$user = Get-ADUser $args[0]
$Acl = Get-Acl $user
$title = New-AccessControlEntry -Principal "Exchange Servers" -ActiveDirectoryRights ReadProperty -ObjectAceType title -AceType Accessdenied -AppliesTo Object
$title = New-AccessControlEntry -Principal "Exchange Servers" -ActiveDirectoryRights ReadProperty -ObjectAceType title -AceType Accessdenied -AppliesTo Object
$telephoneNumber = New-AccessControlEntry -Principal "Exchange Servers" -ActiveDirectoryRights ReadProperty -ObjectAceType telephoneNumber -AceType Accessdenied -AppliesTo Object
$Acl | Add-AccessControlEntry -AceObject $title
$Acl | Add-AccessControlEntry -AceObject $telephoneNumber
$Acl | Set-Acl