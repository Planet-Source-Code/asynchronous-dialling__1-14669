VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "RAS Asynchronous Dialing"
   ClientHeight    =   4440
   ClientLeft      =   1260
   ClientTop       =   2076
   ClientWidth     =   5592
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4440
   ScaleWidth      =   5592
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   4884
      Top             =   3744
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Property"
      Height          =   516
      Left            =   3372
      TabIndex        =   12
      Top             =   3720
      Width           =   1236
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   912
      Left            =   624
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   2256
      Width           =   4404
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connection Status"
      Height          =   1488
      Left            =   240
      TabIndex        =   10
      Top             =   1860
      Width           =   5148
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   516
      Left            =   1764
      TabIndex        =   9
      Top             =   3732
      Width           =   1236
   End
   Begin VB.ComboBox cboConnection 
      Height          =   288
      Left            =   60
      TabIndex        =   8
      Top             =   492
      Width           =   3072
   End
   Begin VB.TextBox txtPhoneNo 
      Height          =   336
      Left            =   3276
      TabIndex        =   4
      Top             =   432
      Width           =   2232
   End
   Begin VB.TextBox txtPassword 
      Height          =   336
      Left            =   3264
      TabIndex        =   3
      Top             =   1128
      Width           =   2232
   End
   Begin VB.TextBox txtUser 
      Height          =   336
      Left            =   84
      TabIndex        =   2
      Top             =   1104
      Width           =   3048
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   528
      Left            =   348
      TabIndex        =   1
      Top             =   3708
      Width           =   1212
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   192
      Left            =   3336
      TabIndex        =   7
      Top             =   912
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Phone No"
      Height          =   192
      Left            =   3312
      TabIndex        =   6
      Top             =   180
      Width           =   744
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "User Name"
      Height          =   192
      Left            =   804
      TabIndex        =   5
      Top             =   852
      Width           =   828
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Dial-Up Connections"
      Height          =   252
      Left            =   60
      TabIndex        =   0
      Top             =   204
      Width           =   3024
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
'
' ===========================================================================
' DATE                              NAME             DESCRIPTION
' --------------------------------  ---------------  -------------
' Wednesday, Jan 24 2001 05:13:1    pramod kumar     module created
'
' ***************************************************************************

Dim MyDialProp As VBRasEntry
Public Sub cmdCancel_Click()
   Dim rtn As Long, lngError As Long
   If g_RASHandle <> 0 Then
       rtn = RasHangUp(g_RASHandle)
       If rtn <> 0 Then
           Debug.Print VBRASErrorHandler(rtn)
       Else
           g_RASHandle = 0
       End If
   Else
       If g_blnDialling = False Then
           Unload Me
       End If
   End If
End Sub

Private Sub cmdConnect_Click()
Dim rtn As Long
   Dim b() As Byte
   Dim myDialParams As VBRasDialParams
   Dim lngHConn As Long
   Dim strPhoneBook As String
If cboConnection.Text = "" Or txtUser = "" Or txtPassword = "" Or txtPhoneNo = "" Then
    MsgBox "Invalid User name or Password or Phone No."
    Exit Sub
End If
sUserName = txtUser
sPhoneNo = txtPhoneNo
If VBRasGetEntryProperties(cboConnection.Text, MyDialProp, vbNullString) = 0 Then
    sDeviceName = MyDialProp.DeviceName
End If
With myDialParams
   .EntryName = cboConnection.Text
   .UserName = txtUser
   .Password = txtPassword
   .PhoneNumber = txtPhoneNo
End With

g_blnDialling = True
bDisconnect = False
cmdConnect.Enabled = False
g_RASHandle = VBAsyncronousDial(vbNullString, cboConnection.Text, myDialParams)
If g_RASHandle <> 0 Then
    Timer1.Enabled = True
End If
End Sub

Private Sub Command1_Click()
RasEditPhonebookEntry Me.hwnd, vbNullString, cboConnection.Text
End Sub

Private Sub Form_Load()
   
   Call Hook(Me.hwnd)
   
   bDisconnect = True
   Timer1.Enabled = False
    
    Dim s As Long, l As Long, ln As Long, a$
    g_RASHandle = 0
    ReDim r(255) As RASENTRYNAME95
    
    r(0).dwSize = 264
    s = 256 * r(0).dwSize
    l = RasEnumEntries(vbNullString, vbNullString, r(0), s, ln)
    For l = 0 To ln - 1
        a$ = StrConv(r(l).szEntryName(), vbUnicode)
        cboConnection.AddItem Left$(a$, InStr(a$, Chr$(0)) - 1)
    Next
    On Local Error Resume Next
    cboConnection.ListIndex = 0
    
    Text1 = "Click Connect to begin connecting. "
End Sub

Private Sub cboConnection_Click()
cboConnection.Text = cboConnection.List(cboConnection.ListIndex)
If VBRasGetEntryProperties(cboConnection.Text, MyDialProp, vbNullString) = 0 Then
    sDeviceName = MyDialProp.DeviceName
End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = g_blnDialling
End Sub
Public Sub Form_Unload(Cancel As Integer)
   UnHook
End Sub


Private Sub Timer1_Timer()
If g_blnDialling = False Then
    cmdConnect.Enabled = True
End If
If bDisconnect = True Then
    Unload Me
End If
End Sub
