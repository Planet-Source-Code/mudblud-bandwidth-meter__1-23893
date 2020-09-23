VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Main 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bandwidth Monitor"
   ClientHeight    =   1215
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   3135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3720
      Top             =   600
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":015C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":02B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Main.frx":0414
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "by Mike15"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.Line Line10 
      X1              =   120
      X2              =   3000
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click to minimize to the system tray"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   850
      Width           =   2895
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bytes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Bytes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   360
      Width           =   495
   End
   Begin VB.Line Line9 
      X1              =   960
      X2              =   960
      Y1              =   360
      Y2              =   840
   End
   Begin VB.Line Line8 
      X1              =   120
      X2              =   3000
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line7 
      X1              =   120
      X2              =   3000
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line6 
      X1              =   3000
      X2              =   3000
      Y1              =   360
      Y2              =   1080
   End
   Begin VB.Line Line5 
      X1              =   1920
      X2              =   3000
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line4 
      X1              =   1920
      X2              =   1920
      Y1              =   120
      Y2              =   360
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   1920
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   1920
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   1080
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Received"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   -120
      TabIndex        =   4
      Top             =   620
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sent"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   -120
      TabIndex        =   3
      Top             =   380
      Width           =   975
   End
   Begin VB.Label lblType 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   130
      Width           =   1815
   End
   Begin VB.Label lblSent 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblRecv 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'constants required by Shell_NotifyIcon API call:
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
Private Const MAXLEN_IFDESCR = 256
Private Const MAXLEN_PHYSADDR = 8
Private Const MAX_INTERFACE_NAME_LEN = 256
Private nid As NOTIFYICONDATA
Private m_objIpHelper As CIpHelper

Private Sub Form_Load()
Set m_objIpHelper = New CIpHelper
With nid
.cbSize = Len(nid)
.hwnd = Me.hwnd
.uId = vbNull
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
.uCallBackMessage = WM_MOUSEMOVE
.hIcon = ImageList1.ListImages(4).Picture
nid.szTip = "Bytes received: " & Trim(Format(m_objIpHelper.BytesReceived, "###,###,###,###")) & vbCrLf & " Bytes sent: " & Trim(Format(m_objIpHelper.BytesSent, "###,###,###,###")) & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub UpdateInterfaceInfo()
Dim objInterface        As CInterface
Static st_objInterface  As CInterface
Static lngBytesRecv     As Long
Static lngBytesSent     As Long
Dim blnIsRecv           As Boolean
Dim blnIsSent           As Boolean
If st_objInterface Is Nothing Then Set st_objInterface = New CInterface
Set objInterface = m_objIpHelper.Interfaces(1)
Select Case objInterface.InterfaceType
Case MIB_IF_TYPE_ETHERNET: lblType.Caption = "Ethernet"
Case MIB_IF_TYPE_FDDI: lblType.Caption = "FDDI"
Case MIB_IF_TYPE_LOOPBACK: lblType.Caption = "Loopback"
Case MIB_IF_TYPE_OTHER: lblType.Caption = "Other"
Case MIB_IF_TYPE_PPP: lblType.Caption = "PPP"
Case MIB_IF_TYPE_SLIP: lblType.Caption = "SLIP"
Case MIB_IF_TYPE_TOKENRING: lblType.Caption = "TokenRing"
End Select
lblRecv.Caption = Trim(Format(m_objIpHelper.BytesReceived, "###,###,###,###"))
lblSent.Caption = Trim(Format(m_objIpHelper.BytesSent, "###,###,###,###"))
Set st_objInterface = objInterface
'---------------
blnIsRecv = (m_objIpHelper.BytesReceived > lngBytesRecv)
blnIsSent = (m_objIpHelper.BytesSent > lngBytesSent)
If blnIsRecv And blnIsSent Then
nid.hIcon = ImageList1.ListImages(4).Picture
ElseIf (Not blnIsRecv) And blnIsSent Then
nid.hIcon = ImageList1.ListImages(3).Picture
ElseIf blnIsRecv And (Not blnIsSent) Then
nid.hIcon = ImageList1.ListImages(2).Picture
ElseIf Not (blnIsRecv And blnIsSent) Then
nid.hIcon = ImageList1.ListImages(1).Picture
End If
lngBytesRecv = m_objIpHelper.BytesReceived
lngBytesSent = m_objIpHelper.BytesSent
nid.szTip = "Bytes received: " & Trim(Format(m_objIpHelper.BytesReceived, "###,###,###,###")) & vbCrLf & " Bytes sent: " & Trim(Format(m_objIpHelper.BytesSent, "###,###,###,###")) & vbNullChar
Shell_NotifyIcon NIM_MODIFY, nid
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = vbBlack
Dim Result As Long
Dim msg As Long
If Me.ScaleMode = vbPixels Then
msg = X
Else
msg = X / Screen.TwipsPerPixelX
End If
Select Case msg
Case WM_LBUTTONDBLCLK    '515 restore form window
If Me.Visible = True Then
Me.Visible = False
Else
Me.Visible = True
Me.SetFocus
End If
End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Form_Terminate()
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = vbBlack
End Sub
Private Sub lblRecv_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = vbBlack
End Sub
Private Sub Label5_Click()
GotoTray
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label5.ForeColor = vbBlue
End Sub

Private Sub Timer1_Timer()
Call UpdateInterfaceInfo
End Sub

