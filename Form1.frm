VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Home Team Editor ML"
   ClientHeight    =   3645
   ClientLeft      =   6630
   ClientTop       =   3000
   ClientWidth     =   4185
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   4185
   Begin VB.Frame Frame3 
      Caption         =   "Casa"
      Height          =   615
      Left            =   2520
      TabIndex        =   5
      Top             =   240
      Width           =   1575
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Equipe 
      Caption         =   "Equipe"
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   2175
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   3975
      Begin VB.Image Image17 
         Height          =   2175
         Left            =   0
         Picture         =   "Form1.frx":000C
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3975
      End
      Begin VB.Image Image16 
         Height          =   2175
         Left            =   0
         Picture         =   "Form1.frx":5BD3
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3975
      End
      Begin VB.Image Image15 
         Height          =   2175
         Left            =   0
         Picture         =   "Form1.frx":10F67
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3975
      End
      Begin VB.Image Image14 
         Height          =   2175
         Left            =   0
         Picture         =   "Form1.frx":1AFB2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3975
      End
      Begin VB.Image Image13 
         Height          =   2175
         Left            =   0
         Picture         =   "Form1.frx":24A8A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3975
      End
      Begin VB.Image Image12 
         Height          =   2175
         Left            =   0
         Picture         =   "Form1.frx":2EF70
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3975
      End
      Begin VB.Image Image11 
         Height          =   2175
         Left            =   0
         Picture         =   "Form1.frx":39165
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3975
      End
      Begin VB.Image Image10 
         Height          =   2175
         Left            =   0
         Picture         =   "Form1.frx":434BA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3975
      End
      Begin VB.Image Image9 
         Height          =   2175
         Left            =   0
         Picture         =   "Form1.frx":4DFD7
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3975
      End
      Begin VB.Image Image8 
         Height          =   2175
         Left            =   0
         Picture         =   "Form1.frx":57B86
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3975
      End
      Begin VB.Image Image7 
         Height          =   2175
         Left            =   0
         Picture         =   "Form1.frx":68445
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3975
      End
      Begin VB.Image Image6 
         Height          =   2175
         Left            =   0
         Picture         =   "Form1.frx":74055
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3975
      End
      Begin VB.Image Image5 
         Height          =   2175
         Left            =   0
         Picture         =   "Form1.frx":808FA
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3975
      End
      Begin VB.Image Image4 
         Height          =   2175
         Left            =   0
         Picture         =   "Form1.frx":8BC18
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3975
      End
      Begin VB.Image Image3 
         Height          =   2175
         Left            =   0
         Picture         =   "Form1.frx":963C6
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3975
      End
      Begin VB.Image Image2 
         Height          =   2175
         Left            =   0
         Picture         =   "Form1.frx":A5D4A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3975
      End
      Begin VB.Image Image1 
         Height          =   2175
         Left            =   0
         Picture         =   "Form1.frx":B03B2
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3975
      End
      Begin VB.Image Image19 
         Height          =   2175
         Left            =   0
         Picture         =   "Form1.frx":BA43A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3975
      End
   End
   Begin VB.TextBox Txendereço 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   5160
      Width           =   3735
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   615
   End
   Begin MSComDlg.CommonDialog Cm1 
      Left            =   0
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salvar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1845
      TabIndex        =   3
      Top             =   3120
      Width           =   645
   End
   Begin VB.Image Imagebtn2 
      Height          =   495
      Left            =   120
      Picture         =   "Form1.frx":C0001
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   3975
   End
   Begin VB.Image Imagebtn1 
      Height          =   495
      Left            =   120
      Picture         =   "Form1.frx":C7B89
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   3975
   End
   Begin VB.Menu About 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
    "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
'=========================================================================
Dim Valor(1 To 32) As Byte


Private Sub About_Click()
Form2.Show vbModal
End Sub

Private Sub Combo1_Click()

Dim rota As String: rota = Txendereço
Combo2.Enabled = True
'Command1.Enabled = True
Imagebtn1.Enabled = True
Imagebtn2.Enabled = True
Image17.Visible = False

Select Case Combo1.Text: Case "Manchester United"
Open rota For Binary As #1: Get #1, 2833127 + 1, Valor(1): Close #1
Text1.Text = Valor(1): End Select

Select Case Combo1.Text: Case "Arsenal"
Open rota For Binary As #1: Get #1, 2833128 + 1, Valor(2): Close #1
Text1.Text = Valor(2): End Select

Select Case Combo1.Text: Case "Chelsea"
Open rota For Binary As #1: Get #1, 2833129 + 1, Valor(3): Close #1
Text1.Text = Valor(3): End Select

Select Case Combo1.Text: Case "Liverpool"
Open rota For Binary As #1: Get #1, 2833130 + 1, Valor(4): Close #1
Text1.Text = Valor(4): End Select

Select Case Combo1.Text: Case "Leeds"
Open rota For Binary As #1: Get #1, 2833131 + 1, Valor(5): Close #1
Text1.Text = Valor(5): End Select

Select Case Combo1.Text: Case "New Castle"
Open rota For Binary As #1: Get #1, 2833132 + 1, Valor(6): Close #1
Text1.Text = Valor(6): End Select

Select Case Combo1.Text: Case "Aston Villa"
Open rota For Binary As #1: Get #1, 2833133 + 1, Valor(7): Close #1
Text1.Text = Valor(7): End Select

Select Case Combo1.Text: Case "Barcelona"
Open rota For Binary As #1: Get #1, 2833134 + 1, Valor(8): Close #1
Text1.Text = Valor(8): End Select

Select Case Combo1.Text: Case "Real Madrid"
Open rota For Binary As #1: Get #1, 2833135 + 1, Valor(9): Close #1
Text1.Text = Valor(9): End Select

Select Case Combo1.Text: Case "Valência"
Open rota For Binary As #1: Get #1, 2833136 + 1, Valor(10): Close #1
Text1.Text = Valor(10): End Select

Select Case Combo1.Text: Case "Deportivo"
Open rota For Binary As #1: Get #1, 2833137 + 1, Valor(11): Close #1
Text1.Text = Valor(11): End Select

Select Case Combo1.Text: Case "Monaco"
Open rota For Binary As #1: Get #1, 2833138 + 1, Valor(12): Close #1
Text1.Text = Valor(12): End Select

Select Case Combo1.Text: Case "Marselha"
Open rota For Binary As #1: Get #1, 2833139 + 1, Valor(13): Close #1
Text1.Text = Valor(13): End Select

Select Case Combo1.Text: Case "PSG"
Open rota For Binary As #1: Get #1, 2833140 + 1, Valor(14): Close #1
Text1.Text = Valor(14): End Select

Select Case Combo1.Text: Case "Bordeaux"
Open rota For Binary As #1: Get #1, 2833141 + 1, Valor(15): Close #1
Text1.Text = Valor(15): End Select

Select Case Combo1.Text: Case "Ajax"
Open rota For Binary As #1: Get #1, 2833142 + 1, Valor(16): Close #1
Text1.Text = Valor(16): End Select

Select Case Combo1.Text: Case "Feyenoord"
Open rota For Binary As #1: Get #1, 2833143 + 1, Valor(17): Close #1
Text1.Text = Valor(17): End Select

Select Case Combo1.Text: Case "PSV"
Open rota For Binary As #1: Get #1, 2833144 + 1, Valor(18): Close #1
Text1.Text = Valor(18): End Select

Select Case Combo1.Text: Case "Internazionale"
Open rota For Binary As #1: Get #1, 2833145 + 1, Valor(19): Close #1
Text1.Text = Valor(19): End Select

Select Case Combo1.Text: Case "Juventus"
Open rota For Binary As #1: Get #1, 2833146 + 1, Valor(20): Close #1
Text1.Text = Valor(20): End Select

Select Case Combo1.Text: Case "Milan"
Open rota For Binary As #1: Get #1, 2833147 + 1, Valor(21): Close #1
Text1.Text = Valor(21): End Select

Select Case Combo1.Text: Case "Lazio"
Open rota For Binary As #1: Get #1, 2833148 + 1, Valor(22): Close #1
Text1.Text = Valor(22): End Select

Select Case Combo1.Text: Case "Parma"
Open rota For Binary As #1: Get #1, 2833149 + 1, Valor(23): Close #1
Text1.Text = Valor(23): End Select

Select Case Combo1.Text: Case "Fiorentina"
Open rota For Binary As #1: Get #1, 2833150 + 1, Valor(24): Close #1
Text1.Text = Valor(24): End Select

Select Case Combo1.Text: Case "Roma"
Open rota For Binary As #1: Get #1, 2833151 + 1, Valor(25): Close #1
Text1.Text = Valor(25): End Select

Select Case Combo1.Text: Case "Borussia Dortmund"
Open rota For Binary As #1: Get #1, 2833152 + 1, Valor(26): Close #1
Text1.Text = Valor(26): End Select

Select Case Combo1.Text: Case "Bayern Munich"
Open rota For Binary As #1: Get #1, 2833153 + 1, Valor(27): Close #1
Text1.Text = Valor(27): End Select

Select Case Combo1.Text: Case "Bayern Leverkusen"
Open rota For Binary As #1: Get #1, 2833154 + 1, Valor(28): Close #1
Text1.Text = Valor(28): End Select

Select Case Combo1.Text: Case "Olympiacos"
Open rota For Binary As #1: Get #1, 2833155 + 1, Valor(29): Close #1
Text1.Text = Valor(29): End Select

Select Case Combo1.Text: Case "Galatasaray"
Open rota For Binary As #1: Get #1, 2833156 + 1, Valor(30): Close #1
Text1.Text = Valor(30): End Select

Select Case Combo1.Text: Case "Dinamo de Kiev"
Open rota For Binary As #1: Get #1, 2833157 + 1, Valor(31): Close #1
Text1.Text = Valor(31): End Select

Select Case Combo1.Text: Case "Boca Juniors"
Open rota For Binary As #1: Get #1, 2833158 + 1, Valor(32): Close #1
Text1.Text = Valor(32): End Select


Select Case Text1.Text: Case Is = "14": Combo2.Text = "GDC_OD": imagens1: End Select
Select Case Text1.Text: Case Is = "15": Combo2.Text = "GDC_VD": imagens2: End Select
Select Case Text1.Text: Case Is = "0": Combo2.Text = "GDC_DD": imagens3: End Select
Select Case Text1.Text: Case Is = "1": Combo2.Text = "GDC_FD": imagens4: End Select
Select Case Text1.Text: Case Is = "2": Combo2.Text = "GDC_JD": imagens5: End Select
Select Case Text1.Text: Case Is = "3": Combo2.Text = "GDC_SD": imagens7: End Select
Select Case Text1.Text: Case Is = "4": Combo2.Text = "GDC_MD": imagens6: End Select
Select Case Text1.Text: Case Is = "5": Combo2.Text = "GDC_AD": imagens8: End Select
Select Case Text1.Text: Case Is = "6": Combo2.Text = "GDC_PD": imagens9: End Select
Select Case Text1.Text: Case Is = "7": Combo2.Text = "GDC_CD": imagens10: End Select
Select Case Text1.Text: Case Is = "8": Combo2.Text = "GDC_HD": imagens11: End Select
Select Case Text1.Text: Case Is = "9": Combo2.Text = "GDC_TD": imagens12: End Select
Select Case Text1.Text: Case Is = "10": Combo2.Text = "GDC_GDJ": imagens13: End Select
Select Case Text1.Text: Case Is = "11": Combo2.Text = "GDC_MDJ": imagens14: End Select
Select Case Text1.Text: Case Is = "12": Combo2.Text = "GDC_RDJ": imagens15: End Select
Select Case Text1.Text: Case Is = "13": Combo2.Text = "GDC_ID": imagens16: End Select
End Sub

Private Sub Combo2_Click()

Select Case Combo2.Text: Case "GDC_OD": Text1.Text = "14": imagens1: End Select
Select Case Combo2.Text: Case "GDC_VD": Text1.Text = "15": imagens2: End Select
Select Case Combo2.Text: Case "GDC_DD": Text1.Text = "0": imagens3: End Select
Select Case Combo2.Text: Case "GDC_FD": Text1.Text = "1": imagens4: End Select
Select Case Combo2.Text: Case "GDC_JD": Text1.Text = "2": imagens5: End Select
Select Case Combo2.Text: Case "GDC_SD": Text1.Text = "3": imagens7: End Select
Select Case Combo2.Text: Case "GDC_MD": Text1.Text = "4": imagens6: End Select
Select Case Combo2.Text: Case "GDC_AD": Text1.Text = "5": imagens8: End Select
Select Case Combo2.Text: Case "GDC_PD": Text1.Text = "6": imagens9: End Select
Select Case Combo2.Text: Case "GDC_CD": Text1.Text = "7": imagens10: End Select
Select Case Combo2.Text: Case "GDC_HD": Text1.Text = "8": imagens11: End Select
Select Case Combo2.Text: Case "GDC_TD": Text1.Text = "9": imagens12: End Select
Select Case Combo2.Text: Case "GDC_GDJ": Text1.Text = "10": imagens13: End Select
Select Case Combo2.Text: Case "GDC_MDJ": Text1.Text = "11": imagens14: End Select
Select Case Combo2.Text: Case "GDC_RDJ": Text1.Text = "12": imagens15: End Select
Select Case Combo2.Text: Case "GDC_ID": Text1.Text = "13": imagens16: End Select
Image17.Visible = False

End Sub





Private Sub Imagebtn1_Click()

Dim rota As String: rota = Txendereço

Select Case Combo1.Text: Case "Manchester United": Valor(1) = Text1.Text
Open rota For Binary As #1: Put #1, 2833127 + 1, Valor(1): Close #1: End Select

Select Case Combo1.Text: Case "Arsenal": Valor(2) = Text1.Text
Open rota For Binary As #1: Put #1, 2833128 + 1, Valor(2): Close #1: End Select

Select Case Combo1.Text: Case "Chelsea": Valor(3) = Text1.Text
Open rota For Binary As #1: Put #1, 2833129 + 1, Valor(3): Close #1: End Select

Select Case Combo1.Text: Case "Liverpool": Valor(4) = Text1.Text
Open rota For Binary As #1: Put #1, 2833130 + 1, Valor(4): Close #1: End Select

Select Case Combo1.Text: Case "Leeds": Valor(5) = Text1.Text
Open rota For Binary As #1: Put #1, 2833131 + 1, Valor(5): Close #1: End Select

Select Case Combo1.Text: Case "New Castle": Valor(6) = Text1.Text
Open rota For Binary As #1: Put #1, 2833132 + 1, Valor(6): Close #1: End Select

Select Case Combo1.Text: Case "Aston Villa": Valor(7) = Text1.Text
Open rota For Binary As #1: Put #1, 2833133 + 1, Valor(7): Close #1: End Select

Select Case Combo1.Text: Case "Barcelona": Valor(8) = Text1.Text
Open rota For Binary As #1: Put #1, 2833134 + 1, Valor(8): Close #1: End Select

Select Case Combo1.Text: Case "Real Madrid": Valor(9) = Text1.Text
Open rota For Binary As #1: Put #1, 2833135 + 1, Valor(9): Close #1: End Select

Select Case Combo1.Text: Case "Valência": Valor(10) = Text1.Text
Open rota For Binary As #1: Put #1, 2833136 + 1, Valor(10): Close #1: End Select

Select Case Combo1.Text: Case "Deportivo": Valor(11) = Text1.Text
Open rota For Binary As #1: Put #1, 2833137 + 1, Valor(11): Close #1: End Select

Select Case Combo1.Text: Case "Monaco": Valor(12) = Text1.Text
Open rota For Binary As #1: Put #1, 2833138 + 1, Valor(12): Close #1: End Select

Select Case Combo1.Text: Case "Marselha": Valor(13) = Text1.Text
Open rota For Binary As #1: Put #1, 2833139 + 1, Valor(13): Close #1: End Select

Select Case Combo1.Text: Case "PSG": Valor(14) = Text1.Text
Open rota For Binary As #1: Put #1, 2833140 + 1, Valor(14): Close #1: End Select

Select Case Combo1.Text: Case "Bordeaux": Valor(15) = Text1.Text
Open rota For Binary As #1: Put #1, 2833141 + 1, Valor(15): Close #1: End Select

Select Case Combo1.Text: Case "Ajax": Valor(16) = Text1.Text
Open rota For Binary As #1: Put #1, 2833142 + 1, Valor(16): Close #1: End Select

Select Case Combo1.Text: Case "Feyenoord": Valor(17) = Text1.Text
Open rota For Binary As #1: Put #1, 2833143 + 1, Valor(17): Close #1: End Select

Select Case Combo1.Text: Case "PSV": Valor(18) = Text1.Text
Open rota For Binary As #1: Put #1, 2833144 + 1, Valor(18): Close #1: End Select

Select Case Combo1.Text: Case "Internazionale": Valor(19) = Text1.Text
Open rota For Binary As #1: Put #1, 2833145 + 1, Valor(19): Close #1: End Select

Select Case Combo1.Text: Case "Juventus": Valor(20) = Text1.Text
Open rota For Binary As #1: Put #1, 2833146 + 1, Valor(20): Close #1: End Select

Select Case Combo1.Text: Case "Milan": Valor(21) = Text1.Text
Open rota For Binary As #1: Put #1, 2833147 + 1, Valor(21): Close #1: End Select

Select Case Combo1.Text: Case "Lazio": Valor(22) = Text1.Text
Open rota For Binary As #1: Put #1, 2833148 + 1, Valor(22): Close #1: End Select

Select Case Combo1.Text: Case "Parma": Valor(23) = Text1.Text
Open rota For Binary As #1: Put #1, 2833149 + 1, Valor(23): Close #1: End Select

Select Case Combo1.Text: Case "Fiorentina": Valor(24) = Text1.Text
Open rota For Binary As #1: Put #1, 2833150 + 1, Valor(24): Close #1: End Select

Select Case Combo1.Text: Case "Roma": Valor(25) = Text1.Text
Open rota For Binary As #1: Put #1, 2833151 + 1, Valor(25): Close #1: End Select

Select Case Combo1.Text: Case "Borussia Dortmund": Valor(26) = Text1.Text
Open rota For Binary As #1: Put #1, 2833152 + 1, Valor(26): Close #1: End Select

Select Case Combo1.Text: Case "Bayern Munich": Valor(27) = Text1.Text
Open rota For Binary As #1: Put #1, 2833153 + 1, Valor(27): Close #1: End Select

Select Case Combo1.Text: Case "Bayern Leverkusen": Valor(28) = Text1.Text
Open rota For Binary As #1: Put #1, 2833154 + 1, Valor(28): Close #1: End Select

Select Case Combo1.Text: Case "Olympiacos": Valor(29) = Text1.Text
Open rota For Binary As #1: Put #1, 2833155 + 1, Valor(29): Close #1: End Select

Select Case Combo1.Text: Case "Galatasaray": Valor(30) = Text1.Text
Open rota For Binary As #1: Put #1, 2833156 + 1, Valor(30): Close #1: End Select

Select Case Combo1.Text: Case "Dinamo de Kiev": Valor(31) = Text1.Text
Open rota For Binary As #1: Put #1, 2833157 + 1, Valor(31): Close #1: End Select

Select Case Combo1.Text: Case "Boca Juniors": Valor(32) = Text1.Text
Open rota For Binary As #1: Put #1, 2833158 + 1, Valor(32): Close #1: End Select

MsgBox "Salvo com sucesso", vbInformation, "Mensagem"
End Sub

Private Sub Form_Load()
Image17.Visible = True
 
Txendereço.Visible = False
Combo2.Enabled = False
'Command1.Enabled = False
Imagebtn1.Enabled = False
Imagebtn2.Enabled = False

'Form1.BackColor = RGB(59, 127, 238)
'Frame1.BackColor = RGB(59, 127, 238)

Cm1.Filter = "We2002 (*.bin)|*.bin": Cm1.ShowOpen
Txendereço.Text = Cm1.FileName

If Txendereço.Text = "" Then
MsgBox "Fim", vbCritical, "Mensagem"
End
End If

With Combo1
.AddItem "Manchester United"
.AddItem "Arsenal"
.AddItem "Chelsea"
.AddItem "Liverpool"
.AddItem "Leeds"
.AddItem "New Castle"
.AddItem "Aston Villa"
.AddItem "Barcelona"
.AddItem "Real Madrid"
.AddItem "Valência"
.AddItem "Deportivo"
.AddItem "Monaco"
.AddItem "Marselha"
.AddItem "PSG"
.AddItem "Bordeaux"
.AddItem "Ajax"
.AddItem "Feyenoord"
.AddItem "PSV"
.AddItem "Internazionale"
.AddItem "Juventus"
.AddItem "Milan"
.AddItem "Lazio"
.AddItem "Parma"
.AddItem "Fiorentina"
.AddItem "Roma"
.AddItem "Borussia Dortmund"
.AddItem "Bayern Munich"
.AddItem "Bayern Leverkusen"
.AddItem "Olympiacos"
.AddItem "Galatasaray"
.AddItem "Dinamo de Kiev"
.AddItem "Boca Juniors"
End With

With Combo2
.AddItem "GDC_OD"
.AddItem "GDC_VD"
.AddItem "GDC_DD"
.AddItem "GDC_FD"
.AddItem "GDC_JD"
.AddItem "GDC_SD"
.AddItem "GDC_MD"
.AddItem "GDC_AD"
.AddItem "GDC_PD"
.AddItem "GDC_CD"
.AddItem "GDC_HD"
.AddItem "GDC_TD"
.AddItem "GDC_GDJ"
.AddItem "GDC_MDJ"
.AddItem "GDC_RDJ"
.AddItem "GDC_ID"
End With

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Imagebtn1.Visible = False
Imagebtn2.Visible = True
End Sub

Private Sub Imagebtn2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Imagebtn1.Visible = True
Imagebtn2.Visible = False
End Sub

'=funções para imprimir imagens na tela==============================================================

Public Sub imagens1()
Image1.Visible = True: Image2.Visible = False
Image3.Visible = False: Image4.Visible = False
Image5.Visible = False: Image6.Visible = False
Image7.Visible = False: Image8.Visible = False
Image9.Visible = False: Image10.Visible = False
Image11.Visible = False: Image12.Visible = False
Image13.Visible = False: Image14.Visible = False
Image15.Visible = False: Image16.Visible = False
Image17.Visible = False
End Sub

Public Sub imagens2()
Image1.Visible = False: Image2.Visible = True
Image3.Visible = False: Image4.Visible = False
Image5.Visible = False: Image6.Visible = False
Image7.Visible = False: Image8.Visible = False
Image9.Visible = False: Image10.Visible = False
Image11.Visible = False: Image12.Visible = False
Image13.Visible = False: Image14.Visible = False
Image17.Visible = False
End Sub

Public Sub imagens3()
Image1.Visible = False: Image2.Visible = False
Image3.Visible = True: Image4.Visible = False
Image5.Visible = False: Image6.Visible = False
Image7.Visible = False: Image8.Visible = False
Image9.Visible = False: Image10.Visible = False
Image11.Visible = False: Image12.Visible = False
Image13.Visible = False: Image14.Visible = False
Image15.Visible = False: Image16.Visible = False
Image17.Visible = False
End Sub

Public Sub imagens4()
Image1.Visible = False: Image2.Visible = False
Image3.Visible = False: Image4.Visible = True
Image5.Visible = False: Image6.Visible = False
Image7.Visible = False: Image8.Visible = False
Image9.Visible = False: Image10.Visible = False
Image11.Visible = False: Image12.Visible = False
Image13.Visible = False: Image14.Visible = False
Image15.Visible = False: Image16.Visible = False
Image17.Visible = False
End Sub

Public Sub imagens5()
Image1.Visible = False: Image2.Visible = False
Image3.Visible = False: Image4.Visible = False
Image5.Visible = True: Image6.Visible = False
Image7.Visible = False: Image8.Visible = False
Image9.Visible = False: Image10.Visible = False
Image11.Visible = False: Image12.Visible = False
Image13.Visible = False: Image14.Visible = False
Image15.Visible = False: Image16.Visible = False
Image17.Visible = False
End Sub

Public Sub imagens6()
Image1.Visible = False: Image2.Visible = False
Image3.Visible = False: Image4.Visible = False
Image5.Visible = False: Image6.Visible = False
Image7.Visible = True: Image8.Visible = False
Image9.Visible = False: Image10.Visible = False
Image11.Visible = False: Image12.Visible = False
Image13.Visible = False: Image14.Visible = False
Image15.Visible = False: Image16.Visible = False
Image17.Visible = False
End Sub

Public Sub imagens7()
Image1.Visible = False: Image2.Visible = False
Image3.Visible = False: Image4.Visible = False
Image5.Visible = False: Image6.Visible = True
Image7.Visible = False: Image8.Visible = False
Image9.Visible = False: Image10.Visible = False
Image11.Visible = False: Image12.Visible = False
Image13.Visible = False: Image14.Visible = False
Image15.Visible = False: Image16.Visible = False
Image17.Visible = False
End Sub

Public Sub imagens8()
Image1.Visible = False: Image2.Visible = False
Image3.Visible = False: Image4.Visible = False
Image5.Visible = False: Image6.Visible = False
Image7.Visible = False: Image8.Visible = True
Image9.Visible = False: Image10.Visible = False
Image11.Visible = False: Image12.Visible = False
Image13.Visible = False: Image14.Visible = False
Image15.Visible = False: Image16.Visible = False
Image17.Visible = False
End Sub

Public Sub imagens9()
Image1.Visible = False: Image2.Visible = False
Image3.Visible = False: Image4.Visible = False
Image5.Visible = False: Image6.Visible = False
Image7.Visible = False: Image8.Visible = False
Image9.Visible = True: Image10.Visible = False
Image11.Visible = False: Image12.Visible = False
Image13.Visible = False: Image14.Visible = False
Image15.Visible = False: Image16.Visible = False
Image17.Visible = False
End Sub

Public Sub imagens10()
Image1.Visible = False: Image2.Visible = False
Image3.Visible = False: Image4.Visible = False
Image5.Visible = False: Image6.Visible = False
Image7.Visible = False: Image8.Visible = False
Image9.Visible = False: Image10.Visible = True
Image11.Visible = False: Image12.Visible = False
Image13.Visible = False: Image14.Visible = False
Image15.Visible = False: Image16.Visible = False
Image17.Visible = False
End Sub

Public Sub imagens11()
Image1.Visible = False: Image2.Visible = False
Image3.Visible = False: Image4.Visible = False
Image5.Visible = False: Image6.Visible = False
Image7.Visible = False: Image8.Visible = False
Image9.Visible = False: Image10.Visible = False
Image11.Visible = True: Image12.Visible = False
Image13.Visible = False: Image14.Visible = False
Image15.Visible = False: Image16.Visible = False
Image17.Visible = False
End Sub

Public Sub imagens12()
Image1.Visible = False: Image2.Visible = False
Image3.Visible = False: Image4.Visible = False
Image5.Visible = False: Image6.Visible = False
Image7.Visible = False: Image8.Visible = False
Image9.Visible = False: Image10.Visible = False
Image11.Visible = False: Image12.Visible = True
Image13.Visible = False: Image14.Visible = False
Image15.Visible = False: Image16.Visible = False
Image17.Visible = False
End Sub

Public Sub imagens13()
Image1.Visible = False: Image2.Visible = False
Image3.Visible = False: Image4.Visible = False
Image5.Visible = False: Image6.Visible = False
Image7.Visible = False: Image8.Visible = False
Image9.Visible = False: Image10.Visible = False
Image11.Visible = False: Image12.Visible = False
Image13.Visible = True: Image14.Visible = False
Image15.Visible = False: Image16.Visible = False
Image17.Visible = False
End Sub

Public Sub imagens14()
Image1.Visible = False: Image2.Visible = False
Image3.Visible = False: Image4.Visible = False
Image5.Visible = False: Image6.Visible = False
Image7.Visible = False: Image8.Visible = False
Image9.Visible = False: Image10.Visible = False
Image11.Visible = False: Image12.Visible = False
Image13.Visible = False: Image14.Visible = True
Image15.Visible = False: Image16.Visible = False
Image17.Visible = False
End Sub

Public Sub imagens15()
Image1.Visible = False: Image2.Visible = False
Image3.Visible = False: Image4.Visible = False
Image5.Visible = False: Image6.Visible = False
Image7.Visible = False: Image8.Visible = False
Image9.Visible = False: Image10.Visible = False
Image11.Visible = False: Image12.Visible = False
Image13.Visible = False: Image14.Visible = False
Image15.Visible = True: Image16.Visible = False
Image17.Visible = False
End Sub

Public Sub imagens16()
Image1.Visible = False: Image2.Visible = True
Image3.Visible = False: Image4.Visible = False
Image5.Visible = False: Image6.Visible = False
Image7.Visible = False: Image8.Visible = False
Image9.Visible = False: Image10.Visible = False
Image11.Visible = False: Image12.Visible = False
Image13.Visible = False: Image14.Visible = False
Image15.Visible = False: Image16.Visible = True
Image17.Visible = False
End Sub



