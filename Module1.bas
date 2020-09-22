Attribute VB_Name = "Module1"
Option Explicit
' ***************************************************************************
'
' Copyright©:    http://www.totalenviro.com/PlatformVB
'                with out Bill McCarthy's help I can't
'                do this simple dialler.Thanx to his help
'                and tutorial.If you doesn't understand
'                functions, type, constants, plz read
'                http://www.totalenviro.com/PlatformVB
'                this will help you a lot
'
' Project:       Dialler
'
' Module:        Form1
'
' Description:   This form contains functions for Asycronous dialing
'                plz, read http://www.totalenviro.com/PlatformVB
'
' ===========================================================================
' DATE                              NAME             DESCRIPTION
' --------------------------------  ---------------  -------------
' Wednesday, Jan 24 2001 05:13:1    pramod kumar     module created
'
' ***************************************************************************

Public Const RAS_MaxDeviceType = 16
Public Const RAS95_MaxDeviceName = 128
Public Const RAS95_MaxEntryName = 256
Public Type RASCONN95
    'set dwsize to 412
    dwSize As Long
    hrasconn As Long
    szEntryName(RAS95_MaxEntryName) As Byte
    szDeviceType(RAS_MaxDeviceType) As Byte
    szDeviceName(RAS95_MaxDeviceName) As Byte
End Type
Public Type RASENTRYNAME95
    'set dwsize to 264
    dwSize As Long
    szEntryName(RAS95_MaxEntryName) As Byte
End Type
Public Type RASDEVINFO
    dwSize As Long
    szDeviceType(RAS_MaxDeviceType) As Byte
    szDeviceName(RAS95_MaxDeviceName) As Byte
End Type
Public Declare Function RasEnumDevices Lib "rasapi32.dll" Alias "RasEnumDevicesA" (lprasdevinfo As Any, lpcb As Long, lpcConnections As Long) As Long
Public Declare Function RasEnumConnections Lib "rasapi32.dll" Alias "RasEnumConnectionsA" (lprasconn As Any, lpcb As Long, lpcConnections As Long) As Long
Public Declare Function RasEnumEntries Lib "rasapi32.dll" Alias "RasEnumEntriesA" (ByVal reserved As String, ByVal lpszPhonebook As String, lprasentryname As Any, lpcb As Long, lpcEntries As Long) As Long


Public Declare Function RasGetEntryDialParams _
      Lib "rasapi32.dll" Alias "RasGetEntryDialParamsA" _
        (ByVal lpszPhonebook As String, _
        lpRasDialParams As Any, _
        blnPasswordRetrieved As Long) As Long

Public Declare Function RasSetEntryDialParams _
      Lib "rasapi32.dll" Alias "RasSetEntryDialParamsA" _
        (ByVal lpszPhonebook As String, _
        lpRasDialParams As Any, _
        ByVal blnRemovePassword As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
        (Destination As Any, Source As Any, ByVal Length As Long)

Public Declare Function RasDial _
      Lib "rasapi32.dll" Alias "RasDialA" _
      (lpRasDialExtensions As Any, _
       ByVal lpszPhonebook As String, _
       lpRasDialParams As Any, _
       ByVal dwNotifierType As Long, _
       ByVal hwndNotifier As Long, _
       lphRasConn As Long) _
As Long
Public Declare Function RasHangUp _
         Lib "rasapi32.dll" Alias "RasHangUpA" _
        (ByVal hrasconn As Long) As Long
   


Public g_blnDialling  As Boolean
'if you close the form when dialling it will crash

Public g_RASHandle As Long
Enum RASCONNSTATE
   RASCS_OpenPort = 0
   RASCS_PortOpened = 1
   RASCS_ConnectDevice = 2
   RASCS_DeviceConnected = 3
   RASCS_AllDevicesConnected = 4
   RASCS_Authenticate = 5
   RASCS_AuthNotify = 6
   RASCS_AuthRetry = 7
   RASCS_AuthCallback = 8
   RASCS_AuthChangePassword = 9
   RASCS_AuthProject = 10
   RASCS_AuthLinkSpeed = 11
   RASCS_AuthAck = 12
   RASCS_ReAuthenticate = 13
   RASCS_Authenticated = 14
   RASCS_PrepareForCallback = 15
   RASCS_WaitForModemReset = 16
   RASCS_WaitForCallback = 17
   RASCS_Projected = 18
   RASCS_StartAuthentication = 19
   RASCS_CallbackComplete = 20
   RASCS_LogonNetwork = 21
   RASCS_SubEntryConnected = 22
   RASCS_SubEntryDisconnected = 23
   RASCS_Interactive = &H1000
   RASCS_RetryAuthentication = &H1001
   RASCS_CallbackSetByCaller = &H1002
   RASCS_PasswordExpired = &H1003
   RASCS_InvokeEapUI = &H1004
   RASCS_connected = &H2000
   RASCS_disconnected = &H2001
End Enum
Public Type RASIPADDR
    a As Byte
    b As Byte
    c As Byte
    d As Byte
End Type

Public Enum RasEntryOptions
   RASEO_UseCountryAndAreaCodes = &H1
   RASEO_SpecificIpAddr = &H2
   RASEO_SpecificNameServers = &H4
   RASEO_IpHeaderCompression = &H8
   RASEO_RemoteDefaultGateway = &H10
   RASEO_DisableLcpExtensions = &H20
   RASEO_TerminalBeforeDial = &H40
   RASEO_TerminalAfterDial = &H80
   RASEO_ModemLights = &H100
   RASEO_SwCompression = &H200
   RASEO_RequireEncryptedPw = &H400
   RASEO_RequireMsEncryptedPw = &H800
   RASEO_RequireDataEncryption = &H1000
   RASEO_NetworkLogon = &H2000
   RASEO_UseLogonCredentials = &H4000
   RASEO_PromoteAlternates = &H8000
   RASEO_SecureLocalFiles = &H10000
   RASEO_RequireEAP = &H20000
   RASEO_RequirePAP = &H40000
   RASEO_RequireSPAP = &H80000
   RASEO_Custom = &H100000
   RASEO_PreviewPhoneNumber = &H200000
   RASEO_SharedPhoneNumbers = &H800000
   RASEO_PreviewUserPw = &H1000000
   RASEO_PreviewDomain = &H2000000
   RASEO_ShowDialingProgress = &H4000000
   RASEO_RequireCHAP = &H8000000
   RASEO_RequireMsCHAP = &H10000000
   RASEO_RequireMsCHAP2 = &H20000000
   RASEO_RequireW95MSCHAP = &H40000000
   RASEO_CustomScript = &H80000000
End Enum

Public Enum RASNetProtocols
   RASNP_NetBEUI = &H1
   RASNP_Ipx = &H2
   RASNP_Ip = &H4
End Enum

Public Enum RasFramingProtocols
   RASFP_Ppp = &H1
   RASFP_Slip = &H2
   RASFP_Ras = &H4
End Enum

Public Type VBRasDialParams
    EntryName As String
    PhoneNumber As String
    CallbackNumber As String
    UserName As String
    Password As String
    Domain As String
    SubEntryIndex As Long
    RasDialFunc2CallbackId As Long
End Type
Public Type VBRasEntry
   Options As RasEntryOptions
   CountryID As Long
   CountryCode As Long
   AreaCode As String
   LocalPhoneNumber As String
   AlternateNumbers As String
   ipAddr As RASIPADDR
   ipAddrDns As RASIPADDR
   ipAddrDnsAlt As RASIPADDR
   ipAddrWins As RASIPADDR
   ipAddrWinsAlt As RASIPADDR
   FrameSize As Long
   fNetProtocols As RASNetProtocols
   FramingProtocol As RasFramingProtocols
   ScriptName As String
   AutodialDll As String
   AutodialFunc As String
   DeviceType As String
   DeviceName As String
   X25PadType As String
   X25Address As String
   X25Facilities As String
   X25UserData As String
   Channels As Long
   NT4En_SubEntries As Long
   NT4En_DialMode As Long
   NT4En_DialExtraPercent As Long
   NT4En_DialExtraSampleSeconds As Long
   NT4En_HangUpExtraPercent As Long
   NT4En_HangUpExtraSampleSeconds As Long
   NT4En_IdleDisconnectSeconds As Long
   Win2000_Type As Long
   Win2000_EncryptionType As Long
   Win2000_CustomAuthKey As Long
   Win2000_guidId(0 To 15) As Byte
   Win2000_CustomDialDll As String
   Win2000_VpnStrategy As Long
End Type
 Public Declare Function RasGetErrorString _
     Lib "rasapi32.dll" Alias "RasGetErrorStringA" _
      (ByVal uErrorValue As Long, ByVal lpszErrorString As String, _
       cBufSize As Long) As Long

Public Declare Function FormatMessage _
     Lib "kernel32" Alias "FormatMessageA" _
      (ByVal dwFlags As Long, lpSource As Any, _
       ByVal dwMessageId As Long, ByVal dwLanguageId As Long, _
       ByVal lpBuffer As String, ByVal nSize As Long, _
       Arguments As Long) As Long
Public Declare Function RasGetEntryProperties _
      Lib "rasapi32.dll" Alias "RasGetEntryPropertiesA" _
       (ByVal lpszPhonebook As String, _
        ByVal lpszEntry As String, _
        lpRasEntry As Any, _
        lpdwEntryInfoSize As Long, _
        lpbDeviceInfo As Any, _
        lpdwDeviceInfoSize As Long) As Long

 
Public Type RASENTRYDIALOG
    dwSize As Long
    hwndOwner As Long
    dwFlags As Long
    xDlg As Long
    yDlg As Long
    szEntry As String
    dwError As Long
    reserved As Long
    reserved2 As Long
End Type
Public Declare Function RasEntryDlg _
      Lib "rasapi32.dll" Alias "RasEntryDlgA" _
       (ByVal lpszPhonebook As String, _
        ByVal lpszEntry As String, _
        lpInfo As RASENTRYDIALOG _
        ) As Long
         
Declare Function RasEditPhonebookEntry Lib "rasapi32.dll" Alias "RasEditPhonebookEntryA" (ByVal hwnd As Long, ByVal lpcstr As String, ByVal lpcstr As String) As Long
Public sDeviceName As String, sPhoneNo As String, sUserName As String
Public bDisconnect As Boolean
Public DebugMsg As String
Function VBRASErrorHandler(rtn As Long) As String
   Dim strError As String, i As Long
   strError = String(512, 0)
   If rtn > 600 Then
      RasGetErrorString rtn, strError, 512&
   Else
      FormatMessage &H1000, ByVal 0&, rtn, 0&, strError, 512, ByVal 0&
   End If
   i = InStr(strError, Chr$(0))
   If i > 1 Then VBRASErrorHandler = Left$(strError, i - 1)
End Function
Function VBRasDialParamsToBytes( _
            udtVBRasDialParamsIN As VBRasDialParams, _
            bytesOut() As Byte) As Boolean
   
   Dim rtn As Long
   Dim blnPsswrd As Long
   Dim b() As Byte
   Dim bLens As Variant
   Dim dwSize As Long, i As Long
   Dim iPos As Long, lngLen As Long
   
   bLens = Array(1060&, 1052&, 816&)
   For i = 0 To 2
      dwSize = bLens(i)
      ReDim b(dwSize - 1)
      CopyMemory b(0), dwSize, 4
      rtn = RasGetEntryDialParams(vbNullString, b(0), blnPsswrd)
      If rtn = 623& Then Exit For
   Next i
   
   If rtn <> 623& Then Exit Function
   
   On Error GoTo badBytes
   ReDim bytesOut(dwSize - 1)
   CopyMemory bytesOut(0), dwSize, 4
   
   If dwSize = 816& Then
      lngLen = 21&
   ElseIf dwSize = 1060& Or dwSize = 1052& Then
      lngLen = 257&
   Else
      'unkown size
      Exit Function
   End If
   iPos = 4
   With udtVBRasDialParamsIN
      CopyStringToByte bytesOut(iPos), .EntryName, lngLen
      iPos = iPos + lngLen: lngLen = 129
      CopyStringToByte bytesOut(iPos), .PhoneNumber, lngLen
      iPos = iPos + lngLen: lngLen = 129
      CopyStringToByte bytesOut(iPos), .CallbackNumber, lngLen
      iPos = iPos + lngLen: lngLen = 257
      CopyStringToByte bytesOut(iPos), .UserName, lngLen
      iPos = iPos + lngLen: lngLen = 257
      CopyStringToByte bytesOut(iPos), .Password, lngLen
      iPos = iPos + lngLen: lngLen = 16
      CopyStringToByte bytesOut(iPos), .Domain, lngLen
      
      If dwSize > 1052& Then
         CopyMemory bytesOut(1052), .SubEntryIndex, 4&
         CopyMemory bytesOut(1056), .RasDialFunc2CallbackId, 4&
      End If
   End With
   VBRasDialParamsToBytes = True
   Exit Function
badBytes:
   'error handling goes here ??
   VBRasDialParamsToBytes = False
End Function
 
Sub CopyStringToByte(bPos As Byte, _
                        strToCopy As String, lngMaxLen As Long)
   Dim lngLen As Long
   lngLen = Len(strToCopy)
   If lngLen = 0 Then
      Exit Sub
   ElseIf lngLen > lngMaxLen Then
      lngLen = lngMaxLen
   End If
   CopyMemory bPos, ByVal strToCopy, lngLen
End Sub
 
 Function VBAsyncronousDial(strPhoneBook As String, _
                              strEntryName As String, ByRef MyParam As VBRasDialParams) As Long
   Dim rtn As Long
   Dim b() As Byte
   Dim lngHConn As Long
   rtn = VBRasDialParamsToBytes(MyParam, b)
   rtn = VBRasSetEntryDialParams(vbNullString, b, True)
   If rtn <> 0 Then GoTo ErrMsg
   'todo: check if rtn = 0 else handle error
   g_blnDialling = True
   rtn = RasDial(ByVal 0&, vbNullString, b(0), _
                     &HFFFFFFFF, Form1.hwnd, lngHConn)
   If rtn <> 0 Then GoTo ErrMsg
   VBAsyncronousDial = lngHConn
   Exit Function
ErrMsg:
MsgBox VBRASErrorHandler(rtn)
End Function
Sub RasDialFunc1(ByVal hrasconn As Long, ByVal unMsg As Long, _
       ByVal rasConnectionState As Long, ByVal dwError As Long, _
       ByVal dwExtendedError As Long)
          
    

    Select Case rasConnectionState
        Case RASCS_OpenPort
            DebugMsg = "The communication port is about to be opened."
        Case RASCS_PortOpened
            DebugMsg = "The communication port has been opened successfully."
        Case RASCS_ConnectDevice
            'DebugMsg = "A device is about to be connected."
            DebugMsg = "Dialing ... " & sPhoneNo & " on " & sDeviceName
        Case RASCS_DeviceConnected
            DebugMsg = "A device has connected successfully"
            
        Case RASCS_AllDevicesConnected
            DebugMsg = "All devices in the device chain have successfully connected. At this point, the physical link is established."
        Case RASCS_Authenticate
            'DebugMsg = "Validating the user name/ password on the specified domain"
            DebugMsg = "Validating the Password for the " & sUserName
        Case RASCS_AuthNotify
            DebugMsg = "An authentication event has occurred"
        Case RASCS_AuthRetry
            DebugMsg = "The client has requested another validation attempt with a new user name/password/domain"
        Case RASCS_AuthCallback
            DebugMsg = "The remote access server has requested a callback number. This occurs only if the user has Set By Caller callback privilege on the server"
        Case RASCS_AuthChangePassword
            DebugMsg = "The client has requested to change the password on the account. "
        Case RASCS_AuthProject
            DebugMsg = "The projection phase is starting."
        Case RASCS_AuthLinkSpeed
            DebugMsg = "The link-speed calculation phase is starting."
        Case RASCS_AuthAck
            DebugMsg = "An authentication request is being acknowledged"
        Case RASCS_ReAuthenticate
            DebugMsg = "Reauthentication (after callback) is starting"
        Case RASCS_Authenticated
            DebugMsg = "The client has successfully completed authentication"
        Case RASCS_PrepareForCallback
            DebugMsg = "The line is about to disconnect in preparation for callback."
        Case RASCS_WaitForModemReset
            DebugMsg = "The client is delaying in order to give the modem time to reset itself in preparation for callback"
        Case RASCS_WaitForCallback
            DebugMsg = "The client is waiting for an incoming call from the remote access server."
        Case RASCS_Projected
            DebugMsg = " This state occurs after the RASCS_AuthProject state. It indicates that projection result information is available. You can access the projection result information by calling  "
        Case RASCS_StartAuthentication
            DebugMsg = "Windows 95 only: Indicates that user authentication is being initiated or retried"
        Case RASCS_CallbackComplete
            DebugMsg = "Windows 95 only: Indicates that the client has been called back and is about to resume authentication."
        Case RASCS_LogonNetwork
            DebugMsg = "Windows 95 only: Indicates that the client is logging on to the network"
        Case RASCS_SubEntryConnected
            DebugMsg = "When dialing a multilink phone-book entry, this state indicates that a subentry has been connected during the dialing process"
        Case RASCS_SubEntryDisconnected
            DebugMsg = "When dialing a multilink phone-book entry, this state indicates that a subentry has been disconnected during the dialing process"
        Case RASCS_Interactive
            DebugMsg = "This state corresponds to the terminal state supported by RASPHONE.EXE."
        Case RASCS_RetryAuthentication
            DebugMsg = "This state corresponds to the retry authentication state supported by RASPHONE.EXE."
        Case RASCS_CallbackSetByCaller
            DebugMsg = "This state corresponds to the callback state supported by RASPHONE.EXE"
        Case RASCS_PasswordExpired
            DebugMsg = "This state corresponds to the change password state supported by RASPHONE.EXE"
        Case RASCS_InvokeEapUI
            DebugMsg = ""
        Case RASCS_connected
            DebugMsg = "Successful connection"
            
        Case RASCS_disconnected
              
              DebugMsg = "Connecting is Cancelled or Failed to Connect." & vbCrLf & "Click Connect to begin connecting again.  To work offline, click Cancel."
               
        Case Else
            DebugMsg = "unknown"
    End Select
    Form1.Text1.Text = Form1.Text1.Text & vbCrLf & DebugMsg
    If rasConnectionState = RASCS_disconnected Then
        g_blnDialling = False
    ElseIf rasConnectionState = RASCS_connected Then
        g_blnDialling = False
        bDisconnect = True
    End If

    If dwError <> 0 Then
      DebugMsg = VBRASErrorHandler(dwError)
      Form1.Text1.Text = Form1.Text1.Text & vbCrLf & DebugMsg
      If g_RASHandle Then Call RasHangUp(g_RASHandle)
      g_blnDialling = False
    End If

End Sub




Function VBRasGetEntryDialParams _
              (bytesOut() As Byte, _
          strPhoneBook As String, strEntryName As String, _
               Optional blnPasswordRetrieved As Boolean) As Long
   
   Dim rtn As Long
   Dim blnPsswrd As Long
   Dim bLens As Variant
   Dim lngLen As Long, i As Long
   
   bLens = Array(1060&, 1052&, 816&)
   'try our three different sizes for RasDialParams
   For i = 0 To 2
      lngLen = bLens(i)
      ReDim bytesOut(lngLen - 1)
      CopyMemory bytesOut(0), lngLen, 4
      If lngLen = 816& Then
         CopyStringToByte bytesOut(4), strEntryName, 20
      Else
         CopyStringToByte bytesOut(4), strEntryName, 256
      End If
      rtn = RasGetEntryDialParams(strPhoneBook, bytesOut(0), blnPsswrd)
      If rtn = 0 Then Exit For
   Next i
   
   blnPasswordRetrieved = blnPsswrd
   VBRasGetEntryDialParams = rtn
End Function
 
Function VBRasSetEntryDialParams _
              (strPhoneBook As String, bytesIn() As Byte, _
               blnRemovePassword As Boolean) As Long
   
   VBRasSetEntryDialParams = RasSetEntryDialParams _
               (strPhoneBook, bytesIn(0), blnRemovePassword)
End Function



'_____________________________________________________________________

Function VBRasGetEntryProperties(strEntryName As String, _
         clsRasEntry As VBRasEntry, _
         Optional strPhoneBook As String) As Long
   
   Dim rtn As Long, lngCb As Long, lngBuffLen As Long
   Dim b() As Byte
   Dim lngPos As Long, lngStrLen As Long

   rtn = RasGetEntryProperties(vbNullString, vbNullString, _
                           ByVal 0&, lngCb, ByVal 0&, ByVal 0&)
   
   rtn = RasGetEntryProperties(strPhoneBook, strEntryName, _
                        ByVal 0&, lngBuffLen, ByVal 0&, ByVal 0&)
   
   If rtn <> 603 Then VBRasGetEntryProperties = rtn: Exit Function
   
   ReDim b(lngBuffLen - 1)
   CopyMemory b(0), lngCb, 4
   
   rtn = RasGetEntryProperties(strPhoneBook, strEntryName, _
                           b(0), lngBuffLen, ByVal 0&, ByVal 0&)
   
   VBRasGetEntryProperties = rtn
   If rtn <> 0 Then Exit Function
   
   CopyMemory clsRasEntry.Options, b(4), 4
   CopyMemory clsRasEntry.CountryID, b(8), 4
   CopyMemory clsRasEntry.CountryCode, b(12), 4
   CopyByteToTrimmedString clsRasEntry.AreaCode, b(16), 11
   CopyByteToTrimmedString clsRasEntry.LocalPhoneNumber, b(27), 129
   
   CopyMemory lngPos, b(156), 4
   If lngPos <> 0 Then
     lngStrLen = lngBuffLen - lngPos
     clsRasEntry.AlternateNumbers = String(lngStrLen, 0)
     CopyMemory ByVal clsRasEntry.AlternateNumbers, _
               b(lngPos), lngStrLen
   End If
   
   CopyMemory clsRasEntry.ipAddr, b(160), 4
   CopyMemory clsRasEntry.ipAddrDns, b(164), 4
   CopyMemory clsRasEntry.ipAddrDnsAlt, b(168), 4
   CopyMemory clsRasEntry.ipAddrWins, b(172), 4
   CopyMemory clsRasEntry.ipAddrWinsAlt, b(176), 4
   CopyMemory clsRasEntry.FrameSize, b(180), 4
   CopyMemory clsRasEntry.fNetProtocols, b(184), 4
   CopyMemory clsRasEntry.FramingProtocol, b(188), 4
   CopyByteToTrimmedString clsRasEntry.ScriptName, b(192), 260
   CopyByteToTrimmedString clsRasEntry.AutodialDll, b(452), 260
   CopyByteToTrimmedString clsRasEntry.AutodialFunc, b(712), 260
   CopyByteToTrimmedString clsRasEntry.DeviceType, b(972), 17
      If lngCb = 1672& Then lngStrLen = 33 Else lngStrLen = 129
   CopyByteToTrimmedString clsRasEntry.DeviceName, b(989), lngStrLen
      lngPos = 989 + lngStrLen
   CopyByteToTrimmedString clsRasEntry.X25PadType, b(lngPos), 33
      lngPos = lngPos + 33
   CopyByteToTrimmedString clsRasEntry.X25Address, b(lngPos), 201
      lngPos = lngPos + 201
   CopyByteToTrimmedString clsRasEntry.X25Facilities, b(lngPos), 201
      lngPos = lngPos + 201
   CopyByteToTrimmedString clsRasEntry.X25UserData, b(lngPos), 201
      lngPos = lngPos + 203
   CopyMemory clsRasEntry.Channels, b(lngPos), 4
   
   If lngCb > 1768 Then 'NT4 Enhancements & Win2000
      CopyMemory clsRasEntry.NT4En_SubEntries, b(1768), 4
      CopyMemory clsRasEntry.NT4En_DialMode, b(1772), 4
      CopyMemory clsRasEntry.NT4En_DialExtraPercent, b(1776), 4
      CopyMemory clsRasEntry.NT4En_DialExtraSampleSeconds, b(1780), 4
      CopyMemory clsRasEntry.NT4En_HangUpExtraPercent, b(1784), 4
      CopyMemory clsRasEntry.NT4En_HangUpExtraSampleSeconds, b(1788), 4
      CopyMemory clsRasEntry.NT4En_IdleDisconnectSeconds, b(1792), 4
      
      If lngCb > 1796 Then ' Win2000
         CopyMemory clsRasEntry.Win2000_Type, b(1796), 4
         CopyMemory clsRasEntry.Win2000_EncryptionType, b(1800), 4
         CopyMemory clsRasEntry.Win2000_CustomAuthKey, b(1804), 4
         CopyMemory clsRasEntry.Win2000_guidId(0), b(1808), 16
         CopyByteToTrimmedString _
                  clsRasEntry.Win2000_CustomDialDll, b(1824), 260
         CopyMemory clsRasEntry.Win2000_VpnStrategy, b(2084), 4
      End If
      
   End If
   
End Function
Sub CopyByteToTrimmedString(strToCopyTo As String, _
                              bPos As Byte, lngMaxLen As Long)
   Dim strTemp As String, lngLen As Long
   strTemp = String(lngMaxLen + 1, 0)
   CopyMemory ByVal strTemp, bPos, lngMaxLen
   lngLen = InStr(strTemp, Chr$(0)) - 1
   strToCopyTo = Left$(strTemp, lngLen)
End Sub
 


