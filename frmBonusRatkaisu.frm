VERSION 5.00
Begin VB.Form frmBonusRatkaisu 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Onnenpyörä - Ratkaisu"
   ClientHeight    =   1500
   ClientLeft      =   2610
   ClientTop       =   3135
   ClientWidth     =   5520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1500
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRatkaisu 
      Height          =   375
      Left            =   120
      MaxLength       =   55
      TabIndex        =   1
      Top             =   360
      Width           =   5295
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Timer tmrAikaB 
      Interval        =   1000
      Left            =   120
      Top             =   1080
   End
   Begin VB.Label lblAika 
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   960
      Width           =   495
   End
   Begin VB.Label lblAikaa 
      Caption         =   "Aikaa jäljellä:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.Label lblRatkaisu 
      AutoSize        =   -1  'True
      Caption         =   "Ratkaisu:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "frmBonusRatkaisu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BonusAikaa As Integer
Private Sub cmdOK_Click()

Dim msg As Integer

If txtRatkaisu.Text = frmKirjaimet.Tehtavateksti Then
    tmrAikaB.Enabled = False
    frmKirjaimet.BonusRatkaise
Else
    txtRatkaisu.Text = ""
End If

End Sub

Private Sub Form_Load()

frmBonusRatkaisu.Top = frmTehtava.Top + 3000
frmBonusRatkaisu.Left = frmTehtava.Left + 200
BonusAikaa = 30
frmBonusRatkaisu.Show
txtRatkaisu.SetFocus

End Sub

Private Sub tmrAikaB_Timer()

Dim msg As Integer

BonusAikaa = BonusAikaa - 1
lblAika.Caption = BonusAikaa

If BonusAikaa < 1 Then
    frmMain.tmrAika.Enabled = False
    frmKirjaimet.tmrAika.Enabled = False
    frmKirjaimet.tmrAika2.Enabled = False
    tmrAikaB.Enabled = False
    msg = MsgBox("Aika loppui.", 0, "Onnenpyörä")
    frmKirjaimet.BonusEiRatkaise
End If

End Sub
