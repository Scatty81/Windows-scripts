' ******************************************************************************
' trustedSite.vbs
' Author:   Rutger Hermarij
' email:    rutger@Scatty.nl
' Date:     07-05-2008
' Version:  1.0
' This script will check add trusted sites for internet explorer,
' if this is performed by a GPO the trusted sites are grayed out for the users.
' ******************************************************************************
Const HKEY_CURRENT_USER = &H80000001	' Hexidecimale value for HKCU
Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}\\.\root\default:StdRegProv")	' Opject to access the registery 


' Call trustedDomain("microsoft.com","http","1") 
' your website -------^-----------^
' http or https ----------------------^---^
' level where the site should be ------------^
' 0 = Internet
' 1 = local intranet
' 2 = trusted site
' 3 = restricted site


Call trustedDomain("microsoft.com","http","1")     ' add the given website as local intranet
Call trustedDomain("bing.com","https","1")    ' add the given website as local intranet


Function trustedDomain(domainName,urlT,level)
    strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\" & domainName
    objReg.CreateKey HKEY_CURRENT_USER,strKeyPath
    objReg.SetDWORDValue HKEY_CURRENT_USER,strKeyPath,urlT,level
End Function
