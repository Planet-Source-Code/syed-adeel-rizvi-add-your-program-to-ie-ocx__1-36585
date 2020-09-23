VERSION 5.00
Begin VB.UserControl AIE 
   ClientHeight    =   885
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1185
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   885
   ScaleWidth      =   1185
   ToolboxBitmap   =   "Main.ctx":0000
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ADD TO Internet Explorer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "AIE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'====================================
' Create By Syed Adeel Hassan Rizvi '
'     Any Question Mail Me          '
'====================================
Public Sub AddTOIE(ByVal Icon1 As String, Icon2 As String, EXEFile As String, ButtonText As String, StatusbarText As String, MenuText As String) ' Add Program To IE
On Error Resume Next
mnuAddIE Icon1, Icon2, EXEFile, ButtonText, StatusbarText, MenuText
End Sub
Public Sub Remove()
mnuDeleteIE ' Remove Program From IE
End Sub
Private Sub UserControl_Initialize()
On Error Resume Next
origWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf AppWndProc)
DetectIE
End Sub
Private Sub UserControl_Resize()
On Error Resume Next
UserControl.Width = 1185
UserControl.Height = 885
End Sub
Private Sub UserControl_Terminate()
On Error Resume Next
SetWindowLong hwnd, GWL_WNDPROC, origWndProc
End Sub
