VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   3120
   ClientLeft      =   7860
   ClientTop       =   2865
   ClientWidth     =   3120
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Sair"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "fabiojraraujo1091@hotmail.com"
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "https://www.facebook.com/groups"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   2160
      Left            =   0
      Picture         =   "Form2.frx":0000
      Top             =   0
      Width           =   3090
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'Private Declare Function ShellExecute Lib "shell32.dll" Alias _
 '   "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
  '  ByVal lpFile As String, ByVal lpParameters As String, _
   ' ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
'=========================================================================

Option Explicit

'Declaraçao para iperlink internet
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWMAXIMIZED = 3
Private Sub Command1_Click()
Unload Me
End Sub

' Open the Web page.
Private Sub Label7_Click()
    ShellExecute ByVal 0&, "open", _
        "https://www.facebook.com/groups/365322350532665/?ref=group_header", _
        vbNullString, vbNullString, _
        SW_SHOWMAXIMIZED
End Sub
 
 
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Label7.Font.Underline = False
End Sub


'O label lblLink tem um ícone de mão vinculado à propriedade MouseIcon
Private Sub Label7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     
     Label7.Font.Underline = True
    
    'Procure no Object Browser (F2) por vbArrowHourglass e verás outros tipos de cursores, mas não o de mão
    'Const vbArrowHourglass = 13
    'Member of VBRUN.MousePointerConstants
    'Arrow and hourglass.
    
    'lblLink.MousePointer = vbArrowHourglass
End Sub




