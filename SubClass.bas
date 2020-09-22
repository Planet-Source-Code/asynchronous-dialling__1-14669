Attribute VB_Name = "SubClass"
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
' Wednesday, Jan 24 2001 05:13:1    pramod kumar     this module for subclassing
'                                                    the form inorder to cancel
'                                                    connection when dialling
'
' ***************************************************************************

Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Const RASDIALEVENT = "RasDialEvent"
Private Const WM_RASDIALEVENT = &HCCCD&
Private m_RasMessage As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Const GWL_WNDPROC = (-4)

Private m_hWnd As Long
Private m_lpPrev As Long


Public Sub Hook(ByVal hwnd As Long)
      If m_hWnd Then Call UnHook
      m_hWnd = hwnd
      m_lpPrev = SetWindowLong(m_hWnd, GWL_WNDPROC, AddressOf WindowProc)
      
      
      If m_RasMessage = 0 Then
         m_RasMessage = RegisterWindowMessage(RASDIALEVENT)
         If m_RasMessage = 0 Then m_RasMessage = WM_RASDIALEVENT
      End If
         
End Sub

Public Sub UnHook()
      If m_hWnd Then
         Call SetWindowLong(m_hWnd, GWL_WNDPROC, m_lpPrev)
         m_hWnd = 0
      End If
End Sub
Private Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   If uMsg = m_RasMessage Then
      Call RasDialFunc1(g_RASHandle, uMsg, wParam, lParam, 0&)
   End If
   WindowProc = CallWindowProc(m_lpPrev, hwnd, uMsg, wParam, lParam)
End Function

