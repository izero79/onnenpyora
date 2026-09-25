VERSION 5.00
Begin VB.Form frmRatkaisu 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Onnenpyörä - Ratkaisu"
   ClientHeight    =   1500
   ClientLeft      =   5505
   ClientTop       =   3270
   ClientWidth     =   5520
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1500
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrAika3 
      Interval        =   1000
      Left            =   840
      Top             =   1080
   End
   Begin VB.Timer tmrAika1 
      Interval        =   1000
      Left            =   120
      Top             =   1080
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtRatkaisu 
      Height          =   375
      Left            =   120
      MaxLength       =   55
      TabIndex        =   0
      Top             =   360
      Width           =   5295
   End
   Begin VB.Label lblRatkaisu 
      AutoSize        =   -1  'True
      Caption         =   "Ratkaisu:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "frmRatkaisu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Aikaa As Integer
Public Kokoaikaa As Integer
Private Sub cmdOK_Click()

Dim msg As Integer

If txtRatkaisu.Text = frmKirjaimet.Tehtavateksti Then
    tmrAika1.Enabled = False
    tmrAika3.Enabled = False
    frmKirjaimet.Ratkaise
    frmRatkaisu.Hide
    Unload frmRatkaisu

Else
    frmTehtava.Visible = False
    frmKirjaimet.Visible = False
    frmRatkaisu.Visible = False
    tmrAika1.Enabled = False
    tmrAika3.Enabled = False
    msg = MsgBox("Valitettavasti se on väärin.", 0, "Onnenpyörä")
    frmTehtava.Visible = True
    frmKirjaimet.Visible = True
    frmMain.VaihdaVuoro
    frmNapit.Enabled = True
    frmNapit.cmdPyoritys.Enabled = True
    frmNapit.cmdOsta.Enabled = True
    frmNapit.cmdRatkaise.Enabled = True
    frmMain.tmrAika.Enabled = True
    frmMain.Aika = 7
    frmRatkaisu.Hide
    Unload frmRatkaisu

End If

End Sub


Private Sub Form_Load()

frmRatkaisu.Top = frmTehtava.Top + 3000
frmRatkaisu.Left = frmTehtava.Left + 200
If frmKirjautuminen.mhCAPS.Value = False Then
    frmKirjautuminen.mhCAPS.Value = True
End If

Aikaa = 7
Kokoaikaa = 45

End Sub



Private Sub tmrAika1_Timer()

Dim msg As Integer

Aikaa = Aikaa - 1

If Aikaa < 1 Then
    frmTehtava.Visible = False
    frmKirjaimet.Visible = False
    frmRatkaisu.Visible = False
    tmrAika1.Enabled = False
    tmrAika3.Enabled = False
    msg = MsgBox("Aika loppui.", 0, "Onnenpyörä")
    frmTehtava.Visible = True
    frmKirjaimet.Visible = True
    frmMain.VaihdaVuoro
    frmNapit.cmdPyoritys.Enabled = True
    frmNapit.cmdOsta.Enabled = True
    frmNapit.cmdRatkaise.Enabled = True
    frmMain.tmrAika.Enabled = True
    frmMain.Aika = 7
    frmRatkaisu.Hide
    Unload frmRatkaisu

End If

End Sub


Private Sub tmrAika3_Timer()

Dim msg As Integer

Kokoaikaa = Kokoaikaa - 1

If Kokoaikaa < 1 Then
    frmTehtava.Visible = False
    frmKirjaimet.Visible = False
    frmRatkaisu.Visible = False
    tmrAika1.Enabled = False
    tmrAika3.Enabled = False
    msg = MsgBox("Aika loppui.", 0, "Onnenpyörä")
    frmTehtava.Visible = True
    frmKirjaimet.Visible = True
    frmMain.VaihdaVuoro
    frmNapit.cmdPyoritys.Enabled = True
    frmNapit.cmdOsta.Enabled = True
    frmNapit.cmdRatkaise.Enabled = True
    frmMain.tmrAika.Enabled = True
    frmMain.Aika = 7
    frmRatkaisu.Hide
    Unload frmRatkaisu

End If


End Sub


Private Sub txtRatkaisu_Change()

Aikaa = 7

End Sub


Private Sub txtRatkaisu_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyBack Or KeyCode = vbKeyDelete Or KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
    txtRatkaisu.Text = ""
End If

End Sub


Private Sub txtRatkaisu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

txtRatkaisu.Text = ""

End Sub


