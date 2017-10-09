Option Explicit
Const  retvalUnknown = 1
Dim    SYSDATA, SYSEXPLANATION  ' Used by Network Monitor, don't change the names


' ///////////////////////////////////////////////////////////////////////////////' // To test a function outside Network Monitor (e.g. using CSCRIPT from the
' // command line), remove the comment character (') in the following 5 lines:
' Dim bResult
' bResult = CheckMcAfeeVirusScanEnterprise8( "localhost", "" )
' WScript.Echo "Return value: [" & bResult & "]"
' WScript.Echo "SYSDATA: [" & SYSDATA & "]"
' WScript.Echo "SYSEXPLANATION: [" & SYSEXPLANATION & "]"' //////////////////////////////////////////////////////////////////////////////


' //////////////////////////////////////////////////////////////////////////////

Function CheckMcAfeeVirusScanEnterprise8( strComputer, strCredentials )

' Description: 
'     Checks if McAfee VirusScan Enterprise is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     CheckMcAfeeVirusScanEnterprise8( "", "" )
' Sample:
'     CheckMcAfeeVirusScanEnterprise8( "localhost", "" )

    Dim objWMIService

    CheckMcAfeeVirusScanEnterprise8 = retvalUnknown  ' Default return value
    SYSDATA                         = ""             ' Not used in this function
    SYSEXPLANATION                  = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckMcAfeeVirusScanEnterprise8 = checkMcAfeeVSEnterprise8WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function


' //////////////////////////////////////////////////////////////////////////////

Function CheckMcAfeeVirusScanPlus2009( strComputer, strCredentials )

' Description: 
'     Checks if McAfee VirusScan Plus 2009 is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     CheckMcAfeeVirusScanPlus2009( "", "" )
' Sample:
'     CheckMcAfeeVirusScanPlus2009( "localhost", "" )

    Dim objWMIService

    CheckMcAfeeVirusScanPlus2009    = retvalUnknown  ' Default return value
    SYSDATA                         = ""             ' Not used in this function
    SYSEXPLANATION                  = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckMcAfeeVirusScanPlus2009 = checkMcAfeeVSPlus2009WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function CheckMcAfeeVirusScanPlus2008( strComputer, strCredentials )

' Description: 
'     Checks if McAfee VirusScan Plus 2008 is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     CheckMcAfeeVirusScanPlus2008( "", "" )
' Sample:
'     CheckMcAfeeVirusScanPlus2008( "localhost", "" )

    Dim objWMIService

    CheckMcAfeeVirusScanPlus2008    = retvalUnknown  ' Default return value
    SYSDATA                         = ""             ' Not used in this function
    SYSEXPLANATION                  = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckMcAfeeVirusScanPlus2008 = checkMcAfeeVSPlus2008WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function CheckNortonInternetSecurity2009( strComputer, strCredentials )

' Description: 
'     Checks if Norton internetSecurity 2009 is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     CheckNortonInternetSecurity2009( "", "" )
' Sample:
'     CheckNortonInternetSecurity2009( "localhost", "" )

    Dim objWMIService

    CheckNortonInternetSecurity2009  = retvalUnknown  ' Default return value
    SYSDATA                          = ""             ' Not used in this function
    SYSEXPLANATION                   = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckNortonInternetSecurity2009  = checkNortonInternetSecurity2009WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function CheckNortonInternetSecurity2008( strComputer, strCredentials )

' Description: 
'     Checks if Norton internetSecurity 2008 is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     CheckNortonInternetSecurity2008( "", "" )
' Sample:
'     CheckNortonInternetSecurity2008( "localhost", "" )

    Dim objWMIService

    CheckNortonInternetSecurity2008  = retvalUnknown  ' Default return value
    SYSDATA                          = ""             ' Not used in this function
    SYSEXPLANATION                   = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckNortonInternetSecurity2008  = checkNortonInternetSecurity2008WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function CheckNortonAntivirus2009( strComputer, strCredentials )

' Description: 
'     Checks if Norton AntiVirus 2009 is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     CheckNortonAntivirus2009( "", "" )
' Sample:
'     CheckNortonAntivirus2009( "localhost", "" )

    Dim objWMIService

    CheckNortonAntivirus2009      = retvalUnknown  ' Default return value
    SYSDATA                       = ""             ' Not used in this function
    SYSEXPLANATION                = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckNortonAntivirus2009      = checkNortonAntivirus2009WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function CheckNortonAntivirus2008( strComputer, strCredentials )

' Description: 
'     Checks if Norton AntiVirus 2008 is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     CheckNortonAntivirus2008( "", "" )
' Sample:
'     CheckNortonAntivirus2008( "localhost", "" )

    Dim objWMIService

    CheckNortonAntivirus2008      = retvalUnknown  ' Default return value
    SYSDATA                       = ""             ' Not used in this function
    SYSEXPLANATION                = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckNortonAntivirus2008      = checkNortonAntivirus2008WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function CheckNortonAntivirus2007( strComputer, strCredentials )

' Description: 
'     Checks if Norton AntiVirus 2007 is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     CheckNortonAntivirus2007( "", "" )
' Sample:
'     CheckNortonAntivirus2007( "localhost", "" )

    Dim objWMIService

    CheckNortonAntivirus2007      = retvalUnknown  ' Default return value
    SYSDATA                       = ""             ' Not used in this function
    SYSEXPLANATION                = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckNortonAntivirus2007      = checkNortonAntivirus2007WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function CheckNortonAntivirus2005( strComputer, strCredentials )

' Description: 
'     Checks if Norton AntiVirus 2005 is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     CheckNortonAntivirus2005( "", "" )
' Sample:
'     CheckNortonAntivirus2005( "localhost", "" )

    Dim objWMIService

    CheckNortonAntivirus2005      = retvalUnknown  ' Default return value
    SYSDATA                       = ""             ' Not used in this function
    SYSEXPLANATION                = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckNortonAntivirus2005      = checkNortonAntivirus2005WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function CheckNOD32v3( strComputer, strCredentials )

' Description: 
'     Checks if NOD32 v3.x is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     CheckNOD32v3( "", "" )
' Sample:
'     CheckNOD32v3( "localhost", "" )

    Dim objWMIService

    CheckNOD32v3      = retvalUnknown  ' Default return value
    SYSDATA           = ""             ' Not used in this function
    SYSEXPLANATION    = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckNOD32v3      = checkNOD32v3WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function CheckNOD32v2( strComputer, strCredentials )

' Description: 
'     Checks if NOD32 v2.x is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     CheckNOD32v2( "", "" )
' Sample:
'     CheckNOD32v2( "localhost", "" )

    Dim objWMIService

    CheckNOD32v2      = retvalUnknown  ' Default return value
    SYSDATA           = ""             ' Not used in this function
    SYSEXPLANATION    = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckNOD32v2      = checkNOD32v2WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function CheckKasperskyInternetSecurity7( strComputer, strCredentials )

' Description: 
'     Checks if Kaspersky internet Security (antivirus, personal firewall, antispam) is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     CheckKasperskyInternetSecurity7( "", "" )
' Sample:
'     CheckKasperskyInternetSecurity7( "localhost", "" )

    Dim objWMIService

    CheckKasperskyInternetSecurity7      = retvalUnknown  ' Default return value
    SYSDATA                              = ""             ' Not used in this function
    SYSEXPLANATION                       = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckKasperskyInternetSecurity7      = checkKasperskyInternetSecurity7WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function


' //////////////////////////////////////////////////////////////////////////////

Function CheckKasperskyAntiVirus7( strComputer, strCredentials )

' Description: 
'     Checks if Kaspersky Anti Virus 7.0 (file antivirus, mail antivirus, web antivirus) is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     CheckKasperskyAntiVirus7( "", "" )
' Sample:
'     CheckKasperskyAntiVirus7( "localhost", "" )

    Dim objWMIService

    CheckKasperskyAntiVirus7             = retvalUnknown  ' Default return value
    SYSDATA                              = ""             ' Not used in this function
    SYSEXPLANATION                       = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckKasperskyAntiVirus7             = checkKasperskyAntiVirus7WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function CheckKasperskyAntivirusServerV5( strComputer, strCredentials )

' Description: 
'     Checks if Kaspersky Antivirus Server is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     CheckKasperskyAntivirusServerV5( "", "" )
' Sample:
'     CheckKasperskyAntivirusServerV5( "localhost", "" )

    Dim objWMIService

    CheckKasperskyAntivirusServerV5      = retvalUnknown  ' Default return value
    SYSDATA                              = ""             ' Not used in this function
    SYSEXPLANATION                       = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckKasperskyAntivirusServerV5      = checkKasperskyAntivirusServerV5WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function


' //////////////////////////////////////////////////////////////////////////////

Function CheckKasperskyAntivirusWksV5( strComputer, strCredentials )

' Description: 
'     Checks if Kaspersky Antivirus Workstation is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     CheckKasperskyAntivirusWksV5( "", "" )
' Sample:
'     CheckKasperskyAntivirusWksV5( "localhost", "" )

    Dim objWMIService

    CheckKasperskyAntivirusWksV5      = retvalUnknown  ' Default return value
    SYSDATA                           = ""             ' Not used in this function
    SYSEXPLANATION                    = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckKasperskyAntivirusWksV5      = checkKasperskyAntivirusWksV5WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function


' //////////////////////////////////////////////////////////////////////////////

Function CheckTrendMicroInternetSecurity2009( strComputer, strCredentials )

' Description: 
'     Checks if Trend Micro Internet Security 2009 is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     CheckTrendMicroInternetSecurity2009( "", "" )
' Sample:
'     CheckTrendMicroInternetSecurity2009( "localhost", "" )

    Dim objWMIService

    CheckTrendMicroInternetSecurity2009 = retvalUnknown  ' Default return value
    SYSDATA                             = ""             ' Not used in this function
    SYSEXPLANATION                      = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckTrendMicroInternetSecurity2009 = checkTrendMicroIS2009WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function CheckTrendMicroInternetSecurity2008( strComputer, strCredentials )

' Description: 
'     Checks if Trend Micro Internet Security 2008 is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     CheckTrendMicroInternetSecurity2008( "", "" )
' Sample:
'     CheckTrendMicroInternetSecurity2008( "localhost", "" )

    Dim objWMIService

    CheckTrendMicroInternetSecurity2008 = retvalUnknown  ' Default return value
    SYSDATA                             = ""             ' Not used in this function
    SYSEXPLANATION                      = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckTrendMicroInternetSecurity2008 = checkTrendMicroIS2008WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function CheckTrendMicroInternetSecurity2007( strComputer, strCredentials )

' Description: 
'     Checks if Trend Micro Internet Security 2007 is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     CheckTrendMicroInternetSecurity2007( "", "" )
' Sample:
'     CheckTrendMicroInternetSecurity2007( "localhost", "" )

    Dim objWMIService

    CheckTrendMicroInternetSecurity2007 = retvalUnknown  ' Default return value
    SYSDATA                             = ""             ' Not used in this function
    SYSEXPLANATION                      = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckTrendMicroInternetSecurity2007 = checkTrendMicroIS2007WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function


' //////////////////////////////////////////////////////////////////////////////

Function CheckAvast4( strComputer, strCredentials )

' Description: 
'     Checks if avast! Antivirus 4 is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     CheckAvast4( "", "" )
' Sample:
'     CheckAvast4( "localhost", "" )

    Dim objWMIService

    CheckAvast4       = retvalUnknown  ' Default return value
    SYSDATA           = ""             ' Not used in this function
    SYSEXPLANATION    = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckAvast4 = checkAvast4WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function


' //////////////////////////////////////////////////////////////////////////////

Function SophosAntiVirus7( strComputer, strCredentials )

' Description: 
'     Checks if Sophos Anti-virus 7 is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     SophosAntiVirus7( "", "" )
' Sample:
'     SophosAntiVirus7( "localhost", "" )

    Dim objWMIService

    SophosAntiVirus7  = retvalUnknown  ' Default return value
    SYSDATA           = ""             ' Not used in this function
    SYSEXPLANATION    = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    SophosAntiVirus7 = checkSophosAntiVirus7WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function


' //////////////////////////////////////////////////////////////////////////////

Function CheckAVGAntiVirus7( strComputer, strCredentials )

' Description: 
'     Checks if AVG AntiVirus 7 is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     CheckAVGAntiVirus7( "", "" )
' Sample:
'     CheckAVGAntiVirus7( "localhost", "" )

    Dim objWMIService

    CheckAVGAntiVirus7 = retvalUnknown  ' Default return value
    SYSDATA            = ""             ' Not used in this function
    SYSEXPLANATION     = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckAVGAntiVirus7 = checkAVGAntiVirus7WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function


' //////////////////////////////////////////////////////////////////////////////

Function CheckNormanAntiVirus5( strComputer, strCredentials )

' Description: 
'     Checks if Norman Anti-Virus 5 is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     CheckNormanAntiVirus5( "", "" )
' Sample:
'     CheckNormanAntiVirus5( "localhost", "" )

    Dim objWMIService

    CheckNormanAntiVirus5      = retvalUnknown  ' Default return value
    SYSDATA                    = ""             ' Not used in this function
    SYSEXPLANATION             = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckNormanAntiVirus5      = checkNormanAntiVirus5WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function


' //////////////////////////////////////////////////////////////////////////////

Function CheckPandaInternetSecurity2009( strComputer, strCredentials )

' Description: 
'     Checks if Panda AntiVirus 2009 is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     CheckPandaInternetSecurity2009( "", "" )
' Sample:
'     CheckPandaInternetSecurity2009( "localhost", "" )

    Dim objWMIService

    CheckPandaInternetSecurity2009    = retvalUnknown  ' Default return value
    SYSDATA                    = ""             ' Not used in this function
    SYSEXPLANATION             = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckPandaInternetSecurity2009    = checkPandaInternetSecurity2009WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function CheckPandaInternetSecurity2008( strComputer, strCredentials )

' Description: 
'     Checks if Panda AntiVirus 2008 is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     CheckPandaInternetSecurity2008( "", "" )
' Sample:
'     CheckPandaInternetSecurity2008( "localhost", "" )

    Dim objWMIService

    CheckPandaInternetSecurity2008    = retvalUnknown  ' Default return value
    SYSDATA                    = ""             ' Not used in this function
    SYSEXPLANATION             = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckPandaInternetSecurity2008    = checkPandaInternetSecurity2008WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function CheckPandaAntiVirus2009( strComputer, strCredentials )

' Description: 
'     Checks if Panda AntiVirus 2009 is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     CheckPandaAntiVirus2009( "", "" )
' Sample:
'     CheckPandaAntiVirus2009( "localhost", "" )

    Dim objWMIService

    CheckPandaAntiVirus2009    = retvalUnknown  ' Default return value
    SYSDATA                    = ""             ' Not used in this function
    SYSEXPLANATION             = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckPandaAntiVirus2009    = checkPandaAntiVirus2009WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function CheckPandaAntiVirus2008( strComputer, strCredentials )

' Description: 
'     Checks if Panda AntiVirus 2008 is running
' Parameters:
'     1) strComputer As String - Hostname or IP address of the computer you want to check
'     2) strCredentials As String - Specify an empty string to use Network Monitor service credentials.
'         To use alternate credentials, enter a server that is defined in Server Credentials table.
'         (To define Server Credentials, choose Tools->Options->Server Credentials)
' Usage:
'     CheckPandaAntiVirus2008( "", "" )
' Sample:
'     CheckPandaAntiVirus2008( "localhost", "" )

    Dim objWMIService

    CheckPandaAntiVirus2008    = retvalUnknown  ' Default return value
    SYSDATA                    = ""             ' Not used in this function
    SYSEXPLANATION             = ""             ' Set initial value

    If( Not getWMIObject( strComputer, strCredentials, objWMIService, SYSEXPLANATION ) ) Then
        Exit Function
    End If

    CheckPandaAntiVirus2008    = checkPandaAntiVirus2008WMI( objWMIService, strComputer, SYSEXPLANATION )
     
End Function








' //////////////////////////////////////////////////////////////////////////////
' //
' // Private Functions
' //   NOTE: Private functions are used by the above functions, and will not
' //         be called directly by the ActiveXperts Network Monitor Service.
' //         Private function names start with a lower case character and will
' //         not be listed in the Network Monitor's function browser.
' //
' //////////////////////////////////////////////////////////////////////////////


Function checkMcAfeeVSEnterprise8WMI( objWMIService, strComputer, BYREF strExplanation )

    Dim lstProcesses, lstServices, result

    checkMcAfeeVSEnterprise8WMI = retvalUnknown  ' Default return value
    strExplanation              = "Unable to check for McAfee on this machine"
    result                      = False

    ' Get the processes list
    If( Not retrieveProcessesList( objWMIService, strComputer, lstProcesses, strExplanation ) ) Then
       Exit Function
    End If   

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If   
 		  			
    result = isProcessRunning( lstProcesses, "McShield.exe", strExplanation )	
    If( result = True ) Then	
       result = isProcessRunning( lstProcesses, "VsTskMgr.exe", strExplanation )
    End If
    If( result = True ) Then
       result = isServiceRunning( lstServices, "McAfeeFramework", "McAfee Framework Service", strExplanation )	
    End If

    If( result = True ) Then 
        checkMcAfeeVSEnterprise8WMI = result
        strExplanation              = "All McAfee VirusScan Enterprise 8 processes and services are running"
    End If
   
    checkMcAfeeVSEnterprise8WMI = result 
End Function

' //////////////////////////////////////////////////////////////////////////////

Function checkMcAfeeVSPlus2009WMI( objWMIService, strComputer, BYREF strExplanation )

    Dim lstProcesses, lstServices, result

    checkMcAfeeVSPlus2009WMI    = retvalUnknown  ' Default return value
    strExplanation              = "Unable to check for McAfee on this machine"
    result                      = False

    ' Get the processes list
    If( Not retrieveProcessesList( objWMIService, strComputer, lstProcesses, strExplanation ) ) Then
       Exit Function
    End If   

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If   
 		  			
    result = isServiceRunning( lstServices, "McShield", "McAfee Real-time Scanner", strExplanation )	
    If( result = True ) Then
       result = isServiceRunning( lstServices, "MpfService", "McAfee Personal Firewall Service", strExplanation )	
    End If
    If( result = True ) Then
       result = isServiceRunning( lstServices, "McProxy", "McAfee Proxy Service", strExplanation )	
    End If
    If( result = True ) Then
       result = isServiceRunning( lstServices, "mcmscsvc", "McAfee Services", strExplanation )	
    End If
    If( result = True ) Then
       result = isServiceRunning( lstServices, "McSysmon", "McAfee SystemGuards", strExplanation )	
    End If


    If( result = True ) Then 
        checkMcAfeeVSPlus2009WMI = result
        strExplanation           = "All McAfee VirusScan Plus 2009 processes and services are running"
    End If
   
    checkMcAfeeVSPlus2009WMI = result 
End Function

' //////////////////////////////////////////////////////////////////////////////

Function checkMcAfeeVSPlus2008WMI( objWMIService, strComputer, BYREF strExplanation )

    Dim lstProcesses, lstServices, result

    checkMcAfeeVSPlus2008WMI    = retvalUnknown  ' Default return value
    strExplanation              = "Unable to check for McAfee on this machine"
    result                      = False

    ' Get the processes list
    If( Not retrieveProcessesList( objWMIService, strComputer, lstProcesses, strExplanation ) ) Then
       Exit Function
    End If   

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If   
 		  			
    result = isServiceRunning( lstServices, "McShield", "McAfee Real-time Scanner", strExplanation )	
    If( result = True ) Then
       result = isServiceRunning( lstServices, "MpfService", "McAfee Personal Firewall Service", strExplanation )	
    End If
    If( result = True ) Then
       result = isServiceRunning( lstServices, "McProxy", "McAfee Proxy Service", strExplanation )	
    End If
    If( result = True ) Then
       result = isServiceRunning( lstServices, "mcmscsvc", "McAfee Services", strExplanation )	
    End If
    If( result = True ) Then
       result = isServiceRunning( lstServices, "McSysmon", "McAfee SystemGuards", strExplanation )	
    End If


    If( result = True ) Then 
        checkMcAfeeVSPlus2008WMI = result
        strExplanation           = "All McAfee VirusScan Plus 2008 processes and services are running"
    End If
   
    checkMcAfeeVSPlus2008WMI = result 
End Function




' //////////////////////////////////////////////////////////////////////////////

Function checkNortonInternetSecurity2009WMI( objWMIService, strComputer, BYREF strExplanation )

    Dim lstServices, lstProcesses, result

    checkNortonInternetSecurity2009WMI  = retvalUnknown  ' Default return value
    strExplanation               = "Unable to check for Norton Antivirus on this machine"
    result                       = False

    ' Get the processes list
    If( Not retrieveProcessesList( objWMIService, strComputer, lstProcesses, strExplanation ) ) Then
       Exit Function
    End If   

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If   
 		  			
    result = isProcessRunning( lstProcesses, "ccSvcHst.exe", strExplanation )	

    ' Manual startup: Symantec Core LC
    ' If( result = True ) Then
    '     result = isServiceRunning( lstServices, "symantec core lc", "Symantec Core LC", strExplanation )
    ' End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "ccevtmgr", "Symantec Event Manager", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "cltnetcnservice", "Symantec Lic NetConnect service", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "ccsetmgr", "Symantec Settings Manager", strExplanation )
    End If
    		
    If result = True Then 
        checkNortonInternetSecurity2009WMI  = True
        strExplanation = "Norton InternetSecurity 2009 is running"
        Exit Function
    End If
 
    checkNortonInternetSecurity2009WMI      = result
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function checkNortonInternetSecurity2008WMI( objWMIService, strComputer, BYREF strExplanation )

    Dim lstServices, lstProcesses, result

    checkNortonInternetSecurity2008WMI  = retvalUnknown  ' Default return value
    strExplanation               = "Unable to check for Norton Antivirus on this machine"
    result                       = False

    ' Get the processes list
    If( Not retrieveProcessesList( objWMIService, strComputer, lstProcesses, strExplanation ) ) Then
       Exit Function
    End If   

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If   
 		  			
    result = isProcessRunning( lstProcesses, "ccSvcHst.exe", strExplanation )	

    ' Manual startup: Symantec Core LC
    ' If( result = True ) Then
    '     result = isServiceRunning( lstServices, "symantec core lc", "Symantec Core LC", strExplanation )
    ' End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "ccevtmgr", "Symantec Event Manager", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "cltnetcnservice", "Symantec Lic NetConnect service", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "ccsetmgr", "Symantec Settings Manager", strExplanation )
    End If
    		
    If result = True Then 
        checkNortonInternetSecurity2008WMI  = True
        strExplanation = "Norton InternetSecurity 2008 is running"
        Exit Function
    End If
 
    checkNortonInternetSecurity2008WMI      = result
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function checkNortonAntivirus2009WMI( objWMIService, strComputer, BYREF strExplanation )

    Dim lstServices, lstProcesses, result

    checkNortonAntivirus2009WMI  = retvalUnknown  ' Default return value
    strExplanation               = "Unable to check for Norton Antivirus on this machine"
    result                       = False

    ' Get the processes list
    If( Not retrieveProcessesList( objWMIService, strComputer, lstProcesses, strExplanation ) ) Then
       Exit Function
    End If   

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If   
 		  			
    result = isProcessRunning( lstProcesses, "ccSvcHst.exe", strExplanation )	

    ' Manual startup: Symantec Core LC
    ' If( result = True ) Then
    '     result = isServiceRunning( lstServices, "symantec core lc", "Symantec Core LC", strExplanation )
    ' End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "ccevtmgr", "Symantec Event Manager", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "cltnetcnservice", "Symantec Lic NetConnect service", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "ccsetmgr", "Symantec Settings Manager", strExplanation )
    End If
    		
    If result = True Then 
        checkNortonAntivirus2009WMI  = True
        strExplanation = "Norton Antivirus 2009 is running"
        Exit Function
    End If
 
    checkNortonAntivirus2009WMI      = result
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function checkNortonAntivirus2008WMI( objWMIService, strComputer, BYREF strExplanation )

    Dim lstServices, lstProcesses, result

    checkNortonAntivirus2008WMI  = retvalUnknown  ' Default return value
    strExplanation               = "Unable to check for Norton Antivirus on this machine"
    result                       = False

    ' Get the processes list
    If( Not retrieveProcessesList( objWMIService, strComputer, lstProcesses, strExplanation ) ) Then
       Exit Function
    End If   

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If   
 		  			
    result = isProcessRunning( lstProcesses, "ccSvcHst.exe", strExplanation )	

    ' Manual startup: Symantec Core LC
    ' If( result = True ) Then
    '     result = isServiceRunning( lstServices, "symantec core lc", "Symantec Core LC", strExplanation )
    ' End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "ccevtmgr", "Symantec Event Manager", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "cltnetcnservice", "Symantec Lic NetConnect service", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "ccsetmgr", "Symantec Settings Manager", strExplanation )
    End If
    		
    If result = True Then 
        checkNortonAntivirus2008WMI  = True
        strExplanation = "Norton Antivirus 2008 is running"
        Exit Function
    End If
 
    checkNortonAntivirus2008WMI      = result
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function checkNortonAntivirus2007WMI( objWMIService, strComputer, BYREF strExplanation )

    Dim lstServices, result

    checkNortonAntivirus2007WMI  = retvalUnknown  ' Default return value
    strExplanation               = "Unable to check for Norton Antivirus on this machine"
    result                       = False

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If   
 		  			
    result = isServiceRunning( lstServices, "symappcore", "Symantec AppCore Service", strExplanation )
    ' Manual startup: Symantec Core LC
    ' If( result = True ) Then
    '    result = isServiceRunning( lstServices, "symantec core lc", "Symantec Core LC", strExplanation )
    ' End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "ccevtmgr", "Symantec Event Manager", strExplanation )
    End If
    ' Manual startup: Symantec IS Password Validation
    ' If( result = True ) Then
    '     result = isServiceRunning( lstServices, "ispwdsvc", "Symantec IS Password Validation", strExplanation )
    ' End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "cltnetcnservice", "Symantec Lic NetConnect service", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "ccsetmgr", "Symantec Settings Manager", strExplanation )
    End If
    		
    If result = True Then 
        checkNortonAntivirus2007WMI  = True
        strExplanation = "Norton Antivirus 2007 is running"
        Exit Function
    End If
 
    checkNortonAntivirus2007WMI      = result
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function checkNortonAntivirus2005WMI( objWMIService, strComputer, BYREF strExplanation )

    Dim lstServices, result

    checkNortonAntivirus2005WMI  = retvalUnknown  ' Default return value
    strExplanation               = "Unable to check for Norton Antivirus on this machine"
    result                       = False

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If 
 		  			
    result = isServiceRunning( lstServices, "navapsvc", "Norton AntiVirus Auto-Protect Service", strExplanation )
    If( result = True ) Then
        result = isServiceRunning( lstServices, "npfmntor", "Norton AntiVirus Firewall Monitor Service", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "ccevtmgr", "Symantec Event Manager", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "ccsetmgr", "Symantec Settings Manager", strExplanation )
    End If
'   If( result = True ) Then
'       result = isServiceRunning( lstServices, "sndsrvc", "Symantec Network Drivers Service", strExplanation )
'   End If
'   If( result = True ) Then
'       result = isServiceRunning( lstServices, "ccpwdsvc", "Symantec Password Validation", strExplanation )
'   End If
'   If( result = True ) Then
'       result = isServiceRunning( lstServices, "spbbcsvc", "Symantec SPBBCSvc", strExplanation )
'   End If
    		
    If result = True Then 
        checkNortonAntivirus2005WMI  = True
        strExplanation = "Norton Antivirus is running"
        Exit Function
    End If
 
    checkNortonAntivirus2005WMI      = result
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function checkNOD32v3WMI( objWMIService, strComputer, BYREF strExplanation )
' NOTE: Some v.3 version have the 'NOD32 Kernel Service' service running,
'       others have 'Eset Service', and it is all v.3. For that reason, we
'       check both services
    Dim lstServices, result

    checkNOD32v3WMI      = retvalUnknown  ' Default return value
    strExplanation       = "Unable to check for NOD32 v3.x on this machine"
    result               = False

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If 
 		  			
    result = isServiceRunning( lstServices, "nod32krn", "NOD32 Kernel Service", strExplanation )	
    If result = True Then 
        checkNOD32v3WMI  = True
        strExplanation   = "NOD32 v3.x is running"
        Exit Function
    End If
    
    result = isServiceRunning( lstServices, "ekrn", "Eset Service", strExplanation )	
    If result = True Then 
        checkNOD32v3WMI  = True
        strExplanation   = "NOD32 v3.x is running"
        Exit Function
    End If
 
    checkNOD32v3WMI      = result
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function checkNOD32v2WMI( objWMIService, strComputer, BYREF strExplanation )

    Dim lstServices, result

    checkNOD32v2WMI      = retvalUnknown  ' Default return value
    strExplanation       = "Unable to check for NOD32 v2.x on this machine"
    result               = False

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If 
 		  			
    result = isServiceRunning( lstServices, "NOD32krn", "NOD32 Kernel Service", strExplanation )		
    If result = True Then 
        checkNOD32v2WMI  = True
        strExplanation   = "NOD32 v2.x is running"
        Exit Function
    End If
 
    checkNOD32v2WMI      = result
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function checkKasperskyInternetSecurity7WMI( objWMIService, strComputer, BYREF strExplanation )

    Dim lstServices, result

    checkKasperskyInternetSecurity7WMI     = retvalUnknown  ' Default return value
    strExplanation                         = "Unable to check for Kaspersky Internet Security on this machine"
    result               		   = False

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If 
 		  			
    result = isServiceRunning( lstServices, "avp", "Kaspersky Anti-Virus Service", strExplanation )		
    If result = True Then 
        checkKasperskyInternetSecurity7WMI = True
        strExplanation                     = "Kaspersky Internet Security is running"
        Exit Function
    End If
 
    checkKasperskyInternetSecurity7WMI     = result
     
End Function


' //////////////////////////////////////////////////////////////////////////////

Function checkKasperskyAntiVirus7WMI( objWMIService, strComputer, BYREF strExplanation )

    Dim lstServices, result

    checkKasperskyAntiVirus7WMI            = retvalUnknown  ' Default return value
    strExplanation                         = "Unable to check for Kaspersky Anti Virus 7 on this machine"
    result               		   = False

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If 
 		  			
    result = isServiceRunning( lstServices, "avp", "Kaspersky Anti-Virus 7.0", strExplanation )		
    If result = True Then 
        checkKasperskyAntiVirus7WMI        = True
        strExplanation                     = "Kaspersky Anti-Virus 7.0 is running"
        Exit Function
    End If
 
    checkKasperskyAntiVirus7WMI            = result
     
End Function


' //////////////////////////////////////////////////////////////////////////////

Function checkKasperskyAntivirusServerV5WMI( objWMIService, strComputer, BYREF strExplanation )

    Dim lstServices, result

    checkKasperskyAntivirusServerV5WMI     = retvalUnknown  ' Default return value
    strExplanation                         = "Unable to check for Kaspersky Anti Virus on this machine"
    result               		   = False

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If 
 		  			
    result = isServiceRunning( lstServices, "klfsblogic", "Kaspersky Anti-Virus Service", strExplanation )		
    If result = True Then 
        checkKasperskyAntivirusServerV5WMI = True
        strExplanation                     = "Kaspersky Anti-Virus Service is running"
        Exit Function
    End If
 
    checkKasperskyAntivirusServerV5WMI     = result
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function checkKasperskyAntivirusWksV5WMI( objWMIService, strComputer, BYREF strExplanation )

    Dim lstServices, result

    checkKasperskyAntivirusWksV5WMI     = retvalUnknown  ' Default return value
    strExplanation                      = "Unable to check for Kaspersky Anti Virus on this machine"
    result               		= False

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If 
 		  			
    result = isServiceRunning( lstServices, "klblmain", "Kaspersky Anti-Virus Service", strExplanation )		
    If result = True Then 
        checkKasperskyAntivirusWksV5WMI = True
        strExplanation                  = "Kaspersky Anti-Virus Service is running"
        Exit Function
    End If
 
    checkKasperskyAntivirusWksV5WMI     = result
     
End Function


' //////////////////////////////////////////////////////////////////////////////

Function checkTrendMicroIS2009WMI( objWMIService, strComputer, BYREF strExplanation )

    Dim lstServices, result

    checkTrendMicroIS2009WMI      = retvalUnknown  ' Default return value
    strExplanation                = "Unable to check for Trend Micro Internet Security 2007 on this machine"
    result                        = False

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If 
 		  			
    result = isServiceRunning( lstServices, "pcctlcom", "Trend Micro Central Control Component", strExplanation )
    If( result = True ) Then
        result = isServiceRunning( lstServices, "tmpfw", "Trend Micro Personal Firewall", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "pcscnsrv", "Trend Micro Protection Against Spyware", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "tmproxy", "Trend Micro Proxy Service", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "tmntsrv", "Trend Micro Real-time Service", strExplanation )
    End If
    		
    If result = True Then 
        checkTrendMicroIS2009WMI  = True
        strExplanation            = "Trend Micro Internet Security 2007 is running"
        Exit Function
    End If
 
    checkTrendMicroIS2009WMI      = result
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function checkTrendMicroIS2008WMI( objWMIService, strComputer, BYREF strExplanation )

    Dim lstServices, result

    checkTrendMicroIS2008WMI      = retvalUnknown  ' Default return value
    strExplanation                = "Unable to check for Trend Micro Internet Security 2007 on this machine"
    result                        = False

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If 
 		  			
    result = isServiceRunning( lstServices, "pcctlcom", "Trend Micro Central Control Component", strExplanation )
    If( result = True ) Then
        result = isServiceRunning( lstServices, "tmpfw", "Trend Micro Personal Firewall", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "pcscnsrv", "Trend Micro Protection Against Spyware", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "tmproxy", "Trend Micro Proxy Service", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "tmntsrv", "Trend Micro Real-time Service", strExplanation )
    End If
    		
    If result = True Then 
        checkTrendMicroIS2008WMI  = True
        strExplanation            = "Trend Micro Internet Security 2007 is running"
        Exit Function
    End If
 
    checkTrendMicroIS2008WMI      = result
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function checkTrendMicroIS2007WMI( objWMIService, strComputer, BYREF strExplanation )

    Dim lstServices, result

    checkTrendMicroIS2007WMI      = retvalUnknown  ' Default return value
    strExplanation                = "Unable to check for Trend Micro Internet Security 2007 on this machine"
    result                        = False

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If 
 		  			
    result = isServiceRunning( lstServices, "pcctlcom", "Trend Micro Central Control Component", strExplanation )
    If( result = True ) Then
        result = isServiceRunning( lstServices, "tmpfw", "Trend Micro Personal Firewall", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "pcscnsrv", "Trend Micro Protection Against Spyware", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "tmproxy", "Trend Micro Proxy Service", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "tmntsrv", "Trend Micro Real-time Service", strExplanation )
    End If
    		
    If result = True Then 
        checkTrendMicroIS2007WMI  = True
        strExplanation            = "Trend Micro Internet Security 2007 is running"
        Exit Function
    End If
 
    checkTrendMicroIS2007WMI      = result
     
End Function



' //////////////////////////////////////////////////////////////////////////////

Function checkSophosAntiVirus7WMI( objWMIService, strComputer, BYREF strExplanation )

    Dim lstServices, result

    checkSophosAntiVirus7WMI     = retvalUnknown  ' Default return value
    strExplanation               = "Unable to check for Sophos Anti-Virus 7 on this machine"
    result                       = False

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If 
 		  			
    result = isServiceRunning( lstServices, "Sophos Agent", "Sophos Agent", strExplanation )
    If( result = True ) Then
        result = isServiceRunning( lstServices, "SAVService", "Sophos Anti-Virus", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "SAVAdminService", "Sophos Anti-Virus status reporter", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "Sophos AutoUpdate Service", "Sophos AutoUpdate Service", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "Sophos Message Router", "Sophos Message Router", strExplanation )
    End If
    		
    If result = True Then 
        checkSophosAntiVirus7WMI  = True
        strExplanation            = "Sophos Anti-Virus 7 is running"
        Exit Function
    End If
 
    checkSophosAntiVirus7WMI      = result
     
End Function


' //////////////////////////////////////////////////////////////////////////////

Function checkAvast4WMI( objWMIService, strComputer, BYREF strExplanation )

    Dim lstServices, result

    checkAvast4WMI      = retvalUnknown  ' Default return value
    strExplanation      = "Unable to check for avast! Antivirus 4 on this machine"
    result              = False

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If 
 		  			
    result = isServiceRunning( lstServices, "avast! Antivirus", "avast! Antivirus", strExplanation )
    If( result = True ) Then
        result = isServiceRunning( lstServices, "aswupdsv", "avast! iAVS4 Control Service", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "avast! Mail scanner", "avast! Mail Scanner", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "avast! Web Scanner", "avast! Web Scanner", strExplanation )
    End If
    		
    If result = True Then 
        checkAvast4WMI  = True
        strExplanation  = "avast! Antivirus 4 is running"
        Exit Function
    End If
 
    checkAvast4WMI      = result
     
End Function



' //////////////////////////////////////////////////////////////////////////////

Function checkAVGAntiVirus7WMI( objWMIService, strComputer, BYREF strExplanation )

    Dim lstServices, result

    checkAVGAntiVirus7WMI      = retvalUnknown  ' Default return value
    strExplanation             = "Unable to check for avast! Antivirus 4 on this machine"
    result                     = False

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If 
 		  			
    result = isServiceRunning( lstServices, "AVGEMS", "AVG E-mail Scanner", strExplanation )
    If( result = True ) Then
        result = isServiceRunning( lstServices, "avg7Alrt", "AVG7 Alert Manager Server", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "avg7Updsvc", "AVG7 Update Service", strExplanation )
    End If

    If result = True Then 
        checkAVGAntiVirus7WMI  = True
        strExplanation         = "AVG AntiVirus 7 is running"
        Exit Function
    End If
 
    checkAVGAntiVirus7WMI      = result
     
End Function


' //////////////////////////////////////////////////////////////////////////////

Function checkNormanAntivirus5WMI( objWMIService, strComputer, BYREF strExplanation )

    Dim lstServices, result

    checkNormanAntivirus5WMI      = retvalUnknown  ' Default return value
    strExplanation                = "Unable to check for Norman Anti Virus on this machine"
    result                        = False

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If 
 		  			
    result = isServiceRunning( lstServices, "norman njeeves", "Norman Njeeves", strExplanation )
    If( result = True ) Then
        result = isServiceRunning( lstServices, "nvcoas", "Norman Virus Control on-access component", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "nvcscheduler", "Norman Virus Control Scheduler", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "norman zanda", "Norman Zanda", strExplanation )
    End If
    		
    If result = True Then 
        checkNormanAntivirus5WMI  = True
        strExplanation            = "Norton Antivirus is running"
        Exit Function
    End If
 
    checkNormanAntivirus5WMI      = result
     
End Function


' //////////////////////////////////////////////////////////////////////////////

Function checkPandaInternetSecurity2009WMI( objWMIService, strComputer, BYREF strExplanation )

    Dim lstServices, lstProcesses, result

    checkPandaInternetSecurity2009WMI     = retvalUnknown  ' Default return value
    strExplanation                 = "Unable to check for Norman Anti Virus on this machine"
    result                         = False

    ' Get the processes list
    If( Not retrieveProcessesList( objWMIService, strComputer, lstProcesses, strExplanation ) ) Then
       Exit Function
    End If  

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If 
 	
    result = isProcessRunning( lstProcesses, "avengine.exe", strExplanation )
    If( result = True ) Then
        result = isServiceRunning( lstServices, "pmshellsrv", "Panda Antispam Engine", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "PAVSRV", "Panda Anti-virus service", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "PAVFNSVR", "Panda Function service", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "PSHost", "Panda Host Service", strExplanation )
    End If		  			
    If( result = True ) Then
        result = isServiceRunning( lstServices, "PSIMSVC", "Panda IManager Service", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "PavPrSrv", "Panda Process Protection Service", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "TPSrv", "Panda TPSrv", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "Panda Software Controller", "Panda Software Controller", strExplanation )
    End If
    		
    If result = True Then 
        checkPandaInternetSecurity2009WMI = True
        strExplanation             = "Panda AntiVirus 2009 is running"
        Exit Function
    End If
 
    checkPandaInternetSecurity2009WMI     = result
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function checkPandaInternetSecurity2008WMI( objWMIService, strComputer, BYREF strExplanation )

    Dim lstServices, lstProcesses, result

    checkPandaInternetSecurity2008WMI     = retvalUnknown  ' Default return value
    strExplanation                 = "Unable to check for Norman Anti Virus on this machine"
    result                         = False

    ' Get the processes list
    If( Not retrieveProcessesList( objWMIService, strComputer, lstProcesses, strExplanation ) ) Then
       Exit Function
    End If  

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If 
 	
    result = isProcessRunning( lstProcesses, "avengine.exe", strExplanation )
    If( result = True ) Then
        result = isServiceRunning( lstServices, "pmshellsrv", "Panda Antispam Engine", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "PAVSRV", "Panda Anti-virus service", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "PAVFNSVR", "Panda Function service", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "PSHost", "Panda Host Service", strExplanation )
    End If		  			
    If( result = True ) Then
        result = isServiceRunning( lstServices, "PSIMSVC", "Panda IManager Service", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "PavPrSrv", "Panda Process Protection Service", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "TPSrv", "Panda TPSrv", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "Panda Software Controller", "Panda Software Controller", strExplanation )
    End If
    		
    If result = True Then 
        checkPandaInternetSecurity2008WMI = True
        strExplanation             = "Panda AntiVirus 2008 is running"
        Exit Function
    End If
 
    checkPandaInternetSecurity2008WMI     = result
     
End Function


' //////////////////////////////////////////////////////////////////////////////

Function checkPandaAntiVirus2009WMI( objWMIService, strComputer, BYREF strExplanation )

    Dim lstServices, lstProcesses, result

    checkPandaAntiVirus2009WMI     = retvalUnknown  ' Default return value
    strExplanation                 = "Unable to check for Norman Anti Virus on this machine"
    result                         = False

    ' Get the processes list
    If( Not retrieveProcessesList( objWMIService, strComputer, lstProcesses, strExplanation ) ) Then
       Exit Function
    End If  

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If 
 	
    result = isProcessRunning( lstProcesses, "avengine.exe", strExplanation )	  			
    If( result = True ) Then
        result = isServiceRunning( lstServices, "PSIMSVC", "Panda IManager Service", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "Panda Software Controller", "Panda Software Controller", strExplanation )
    End If
    		
    If result = True Then 
        checkPandaAntiVirus2009WMI = True
        strExplanation             = "Panda AntiVirus 2009 is running"
        Exit Function
    End If
 
    checkPandaAntiVirus2009WMI     = result
     
End Function

' //////////////////////////////////////////////////////////////////////////////

Function checkPandaAntiVirus2008WMI( objWMIService, strComputer, BYREF strExplanation )

    Dim lstServices, lstProcesses, result

    checkPandaAntiVirus2008WMI     = retvalUnknown  ' Default return value
    strExplanation                 = "Unable to check for Norman Anti Virus on this machine"
    result                         = False

    ' Get the processes list
    If( Not retrieveProcessesList( objWMIService, strComputer, lstProcesses, strExplanation ) ) Then
       Exit Function
    End If  

    ' Get the services list
    If( Not retrieveServicesList( objWMIService, strComputer, lstServices, strExplanation ) ) Then
       Exit Function
    End If 
 	
    result = isProcessRunning( lstProcesses, "avengine.exe", strExplanation )	  			
    If( result = True ) Then
        result = isServiceRunning( lstServices, "PSIMSVC", "Panda IManager Service", strExplanation )
    End If
    If( result = True ) Then
        result = isServiceRunning( lstServices, "Panda Software Controller", "Panda Software Controller", strExplanation )
    End If
    		
    If result = True Then 
        checkPandaAntiVirus2008WMI = True
        strExplanation             = "Panda AntiVirus 2008 is running"
        Exit Function
    End If
 
    checkPandaAntiVirus2008WMI     = result
     
End Function






' //////////////////////////////////////////////////////////////////////////////

Function retrieveProcessesList( objWMIService, strComputer, BYREF lstProcesses, BYREF strSysExplanation )
' Retrieve the list of running services	

    retrieveProcessesList = False
    Set lstProcesses      = Nothing

On Error Resume Next

    Set lstProcesses      = objWMIService.ExecQuery( "Select * from Win32_Process" )  
    If( Err.Number <> 0 ) Then
        strSysExplanation  = "Unable to query WMI on computer [" & strComputer & "]"
        Exit Function
    End If
    If( lstProcesses.Count <= 0  ) Then
        strSysExplanation  = "Win32_Process class does not exist on computer [" & strComputer & "]"
        Exit Function
    End If 

On Error Goto 0

    retrieveProcessesList = True    

End Function

' //////////////////////////////////////////////////////////////////////////////

Function retrieveServicesList( objWMIService, strComputer, BYREF lstServices, BYREF strSysExplanation )
' Retrieve the list of running services	

    retrieveServicesList = False
    Set lstServices      = Nothing

On Error Resume Next

    Set lstServices      = objWMIService.ExecQuery( "Select * from Win32_Service WHERE state = ""Running""" )
    If( Err.Number <> 0 ) Then
        strSysExplanation  = "Unable to query WMI on computer [" & strComputer & "]"
        Exit Function
    End If
    If( lstServices.Count <= 0  ) Then
        strSysExplanation  = "Win32_Service class does not exist on computer [" & strComputer & "]"
        Exit Function
    End If 

    retrieveServicesList = True    

End Function


' //////////////////////////////////////////////////////////////////////////////

Function isProcessRunning( BYREF lstProcesses, strProcess, BYREF strExplanation )
' Check if a given service exists as running service in the services list
On Error Resume Next
    
    Dim objProcess
    				
    For Each objProcess in lstProcesses			
	
        If( Err.Number <> 0 ) Then
            isProcessRunning  = retvalUnknown
            strExplanation    = "Unable to list processes" 
            Exit Function
        End If	 
	   
        ' Check If this is the service we are looking for
        If( LCase( objProcess.Name ) = LCase( strProcess ) ) Then				
            isProcessRunning  = True
            Exit Function
        End If
    Next
    
    ' The process was not found, show an error message
    strExplanation            = "'" & strProcess & "' process is not running"    
    isProcessRunning          = False

End Function


' //////////////////////////////////////////////////////////////////////////////

Function isServiceRunning( BYREF lstServices, strServiceName, strServiceDescription, BYREF strExplanation )
' Check if a given service exists as running service in the services list
On Error Resume Next
    
    Dim objService
    				
    For Each objService in lstServices			
		
        If( Err.Number <> 0 ) Then
            isServiceRunning = retvalUnknown
            strExplanation      = "Unable to list services" 
            Exit Function
        End If	 
	   
        ' Check If this is the service we are looking for
        If( LCase( objService.Name ) = LCase( strServiceName ) ) Then				
            isServiceRunning = True
            Exit Function
        End If
    Next
    
    ' The service was not found, show an error message
    strExplanation              = "'" & strServiceDescription & "' service is not running"    
    isServiceRunning         = False

End Function


' //////////////////////////////////////////////////////////////////////////////

Function getWMIObject( strComputer, strCredentials, BYREF objWMIService, BYREF strSysExplanation )	

On Error Resume Next

    Dim objNMServerCredentials, objSWbemLocator, colItems
    Dim strUsername, strPassword

    getWMIObject              = False

    Set objWMIService         = Nothing
    
    If( strCredentials = "" ) Then	
        ' Connect to remote host on same domain using same security context
        Set objWMIService     = GetObject( "winmgmts:{impersonationLevel=Impersonate}!\\" & strComputer &"\root\cimv2" )
    Else
        Set objNMServerCredentials = CreateObject( "ActiveXperts.NMServerCredentials" )

        strUsername           = objNMServerCredentials.GetLogin( strCredentials )
        strPassword           = objNMServerCredentials.GetPassword( strCredentials )

        If( strUsername = "" ) Then
            getWMIObject      = False
            strSysExplanation = "No alternate credentials defined for [" & strCredentials & "]. In the Manager application, select 'Options' from the 'Tools' menu and select the 'Server Credentials' tab to enter alternate credentials"
            Exit Function
        End If
	
        ' Connect to remote host using different security context and/or different domain 
        Set objSWbemLocator   = CreateObject( "WbemScripting.SWbemLocator" )
        Set objWMIService     = objSWbemLocator.ConnectServer( strComputer, "root\cimv2", strUsername, strPassword )

        If( Err.Number <> 0 ) Then
            objWMIService     = Nothing
            getWMIObject      = False
            strSysExplanation = "Unable to access [" & strComputer & "]. Possible reasons: WMI not running on the remote server, Windows firewall is blocking WMI calls, insufficient rights, or remote server down"
            Exit Function
        End If

        objWMIService.Security_.ImpersonationLevel = 3

    End If
	
    If( Err.Number <> 0 ) Then
        objWMIService         = Nothing
        getWMIObject          = False
        strSysExplanation     = "Unable to access '" & strComputer & "'. Possible reasons: no WMI installed on the remote server, no rights to access remote WMI service, or remote server down"
        Exit Function
    End If    

    getWMIObject              = True 

End Function