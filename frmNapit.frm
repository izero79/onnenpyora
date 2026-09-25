VERSION 5.00
Begin VB.Form frmNapit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Onnenpyörä"
   ClientHeight    =   2220
   ClientLeft      =   360
   ClientTop       =   930
   ClientWidth     =   2220
   Icon            =   "frmNapit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2220
   ScaleWidth      =   2220
   Begin VB.CommandButton cmdTakaisin 
      Caption         =   "&Takaisin peliin"
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   1680
      Width           =   375
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdRatkaise 
      Caption         =   "&Ratkaise"
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdOsta 
      Caption         =   "Osta &vokaali"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdPyoritys 
      Caption         =   "&Pyöritä"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   1215
   End
   Begin VB.Menu mnuPeli 
      Caption         =   "&Peli"
      Begin VB.Menu mnuUusipeli 
         Caption         =   "&Uusi peli"
      End
      Begin VB.Menu mnuRekisterointi 
         Caption         =   "&Rekisteröinti"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuLopeta 
         Caption         =   "&Lopeta"
      End
   End
   Begin VB.Menu mnuOhje 
      Caption         =   "&Ohje"
      Begin VB.Menu mnuOhjeSisallys 
         Caption         =   "&Ohjeen sisällys"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuViiva1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInternet 
         Caption         =   "T&ESI software internetissä"
      End
      Begin VB.Menu mnuLaheta 
         Caption         =   "&Lähetä sähköpostia"
      End
      Begin VB.Menu mnuViiva2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTietoja 
         Caption         =   "&Tietoja Onnenpyörästä"
      End
   End
End
Attribute VB_Name = "frmNapit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private pienennetty As Integer



Public Sub Uusipeli()

frmMain.Aika = 0
frmMain.YllatysSaatu = 0
frmMain.SYllatysSaatu = 0
frmMain.YllatyksenVuoro = 0
frmMain.Sektoriteksti = ""
frmMain.Kierros = 0
frmMain.KokoPisteetP1 = 0
frmMain.KokoPisteetP2 = 0
frmMain.KokoPisteetP3 = 0
frmMain.PisteetP1 = 0
frmMain.PisteetP2 = 0
frmMain.PisteetP3 = 0
frmMain.RahaMaaraP1 = 0
frmMain.RahaMaaraP2 = 0
frmMain.RahaMaaraP3 = 0
frmMain.Arvo = 0
frmMain.Tehtava1 = 0
frmMain.Tehtava2 = 0
frmMain.Tehtava3 = 0
frmMain.Tehtava4 = 0
frmMain.Tehtava5 = 0
frmMain.UusiTehtava = 0
frmMain.Pelaaja = 0
frmMain.SektoriPisteet = 0
frmMain.Vokaaliostettu = 0
frmMain.Kelaus = 0
frmMain.numero = 0
frmMain.maxnumero = 0
frmMain.maxtnumero = 0
frmMain.Palkinnonaika = 0
frmMain.Palkinto = 0
frmMain.SPalkinnonsaaja = 0
frmMain.PalkintoVari = 0
frmMain.BPalkinnonnumero1 = 0
frmMain.BPalkinnonnumero2 = 0
frmMain.BPalkinnonnumero3 = 0
frmMain.BPalkinnonteksti1 = ""
frmMain.BPalkinnonteksti2 = ""
frmMain.BPalkinnonteksti3 = ""
frmMain.TodellaLoppu = 0

frmBonusRatkaisu.BonusAikaa = 0

frmBPalkinnot.Palkintosarja = 0

frmKirjaimet.KAika = 0
frmKirjaimet.Tehtavateksti = ""
frmKirjaimet.Tehtavarivi = ""
frmKirjaimet.Teksti = ""
frmKirjaimet.Teksti2 = ""
frmKirjaimet.Teksti3 = ""
frmKirjaimet.Teksti4 = ""
frmKirjaimet.k1 = ""
frmKirjaimet.k2 = ""
frmKirjaimet.k3 = ""
frmKirjaimet.k4 = ""
frmKirjaimet.k5 = ""
frmKirjaimet.k6 = ""
frmKirjaimet.k7 = ""
frmKirjaimet.k8 = ""
frmKirjaimet.k9 = ""
frmKirjaimet.k10 = ""
frmKirjaimet.k11 = ""
frmKirjaimet.k12 = ""
frmKirjaimet.k13 = ""
frmKirjaimet.Apu1 = ""
frmKirjaimet.Apu2 = ""
frmKirjaimet.Apu3 = ""
frmKirjaimet.SB = 0
frmKirjaimet.Valittujakirjaimia = 0
frmKirjaimet.Bonus = 0
frmKirjaimet.Valiaika = 0
frmKirjaimet.BSanottu = 0
frmKirjaimet.KSanottu = 0
frmKirjaimet.A = 0
frmKirjaimet.B = 0
frmKirjaimet.C = 0
frmKirjaimet.D = 0
frmKirjaimet.E = 0
frmKirjaimet.F = 0
frmKirjaimet.G = 0
frmKirjaimet.H = 0
frmKirjaimet.II = 0
frmKirjaimet.J = 0
frmKirjaimet.K = 0
frmKirjaimet.L = 0
frmKirjaimet.M = 0
frmKirjaimet.N = 0
frmKirjaimet.O = 0
frmKirjaimet.P = 0
frmKirjaimet.Q = 0
frmKirjaimet.R = 0
frmKirjaimet.S = 0
frmKirjaimet.T = 0
frmKirjaimet.U = 0
frmKirjaimet.V = 0
frmKirjaimet.W = 0
frmKirjaimet.X = 0
frmKirjaimet.Y = 0
frmKirjaimet.Z = 0
frmKirjaimet.AO = 0
frmKirjaimet.AE = 0
frmKirjaimet.OE = 0

frmKirjautuminen.Nimi1 = ""
frmKirjautuminen.Nimi2 = ""
frmKirjautuminen.Nimi3 = ""
frmKirjautuminen.Pluku = 0

frmLoppupisteet.BonukseenMenija = 0
frmLoppupisteet.SuperBonus = 0
frmLoppupisteet.Arvonta = 0
frmLoppupisteet.SBonus = ""

frmPalkinto.Palkintonumero = 0
frmPalkinto.Palkinto = ""

frmRatkaisu.Aikaa = 0
frmRatkaisu.Kokoaikaa = 0

frmMain.TodellaLoppu = 0
Unload frmTehtava
Unload frmBPalkinnot
Unload frmPalkinto
Unload frmBonusRatkaisu
Unload frmKirjaimet
Unload frmKirjautuminen
Unload frmLoppupisteet
Unload frmMain
Unload frmNapit
Unload frmRatkaisu
Load frmKirjautuminen
frmKirjautuminen.Show

End Sub

Private Sub cmdOsta_Click()

frmMain.Aika = 7
frmMain.Vokaaliostettu = 1

Select Case frmMain.Pelaaja

Case 1
    
    If frmMain.PisteetP1 >= 300 Then
        frmMain.VokaalitNakyviin
        frmNapit.cmdPyoritys.Enabled = False
        frmNapit.cmdOsta.Enabled = False
        frmNapit.cmdRatkaise.Enabled = False
        frmKirjaimet.Enabled = True
        Exit Sub
    
    End If

Case 2
    
    If frmMain.PisteetP2 >= 300 Then
        frmMain.VokaalitNakyviin
        frmNapit.cmdPyoritys.Enabled = False
        frmNapit.cmdOsta.Enabled = False
        frmNapit.cmdRatkaise.Enabled = False
        frmKirjaimet.Enabled = True
        Exit Sub
    
    End If

Case 3
    
    If frmMain.PisteetP3 >= 300 Then
        frmMain.VokaalitNakyviin
        frmNapit.cmdPyoritys.Enabled = False
        frmNapit.cmdOsta.Enabled = False
        frmNapit.cmdRatkaise.Enabled = False
        frmKirjaimet.Enabled = True
        Exit Sub
    
    End If

End Select

Dim msg As Integer

frmTehtava.Visible = False
frmKirjaimet.Visible = False
frmMain.tmrAika.Enabled = False
msg = MsgBox("Sinulla ei ole tarpeeksi pisteitä.", 0, "Onnenpyörä")
frmMain.tmrAika.Enabled = True
frmTehtava.Visible = True
frmKirjaimet.Visible = True

End Sub

Private Sub cmdPyoritys_Click()

frmMain.tmrAika.Enabled = False
frmMain.Aika = 7
cmdOsta.Enabled = True
frmNapit.cmdPyoritys.Enabled = False
frmNapit.cmdOsta.Enabled = False
frmNapit.cmdRatkaise.Enabled = False
frmMain.Pyoraytys

End Sub



Private Sub cmdRatkaise_Click()

Dim ratkaisu As String

Load frmRatkaisu
frmRatkaisu.Show
frmNapit.cmdPyoritys.Enabled = False
frmNapit.cmdOsta.Enabled = False
frmNapit.cmdRatkaise.Enabled = False
frmNapit.Enabled = False
frmMain.tmrAika.Enabled = False

End Sub



Public Sub cmdTakaisin_Click()

cmdTakaisin.Enabled = False
cmdPyoritys.Enabled = True
cmdOsta.Enabled = True
cmdRatkaise.Enabled = True
frmKirjaimet.Show
frmTehtava.Show
frmMain.tmrAika.Enabled = True
frmKirjaimet.tmrAika.Enabled = True
frmKirjaimet.tmrAika2.Enabled = True
frmKirjautuminen.mhCAPS.Value = True
frmNapit.SetFocus


End Sub

Private Sub Command3_Click()

frmKirjaimet.Ratkaise

End Sub


Private Sub Form_Load()

Select Case Screen.Width / 15
Case 640
    frmNapit.Top = 0
    frmNapit.Left = 0
Case 800
    frmNapit.Top = 100
    frmNapit.Left = 100
Case Else
    frmNapit.Top = 675
    frmNapit.Left = 1500
End Select

pienennetty = 0

If frmRekisterointi.Rekisteroity = 1 Then
    mnuRekisterointi.Visible = False
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If frmMain.TodellaLoppu = 1 Then
    
    frmTehtava.Visible = False
    frmKirjaimet.Visible = False
    frmMain.tmrAika.Enabled = False
    msg = MsgBox("Haluatko varmasti lopettaa?", 4, "Onnenpyörä")
    
    If msg = 6 Then
        frmKirjautuminen.mhCAPS.Value = False
        End
    Else
'        frmTehtava.Visible = True
'        frmKirjaimet.Visible = True
'        frmMain.tmrAika.Enabled = True
        Cancel = 1
    End If

End If

End Sub

Private Sub Form_Resize()

If frmNapit.WindowState = 1 Then
    frmTehtava.Enabled = False
    frmMain.Enabled = False
    frmKirjaimet.Enabled = False
    frmMain.tmrAika.Enabled = False
    frmTehtava.Hide
    frmMain.Hide
    frmKirjaimet.Hide
    cmdTakaisin.Enabled = True
    cmdPyoritys.Enabled = False
    cmdOsta.Enabled = False
    cmdRatkaise.Enabled = False
    frmKirjaimet.tmrAika.Enabled = False
    frmKirjaimet.tmrAika2.Enabled = False
    frmKirjautuminen.mhCAPS.Value = False
    pienennetty = 1
End If

If frmNapit.WindowState = 0 And pienennetty = 0 Then
    frmTehtava.Enabled = True
    frmMain.Enabled = True
    frmKirjaimet.Enabled = True
    frmMain.tmrAika.Enabled = True
    frmTehtava.Show
    frmMain.Show
    frmKirjaimet.Show
    pienennetty = 0
End If

If frmNapit.WindowState = 0 And pienennetty = 1 Then
    frmTehtava.Enabled = True
    frmMain.Enabled = True
    frmKirjaimet.Enabled = True
    frmMain.Show
    frmNapit.Show
    pienennetty = 0
End If
    
End Sub


Public Sub mnuInternet_Click()

Dim i As Long

pathi = App.Path
i = ShellExecute(Me.hWnd, vbNullString, "http://members.xoom.com/tesisoftware", vbNullString, pathi, SW_SHOWNORMAL)

End Sub

Public Sub mnuLaheta_Click()

Dim i As Long

pathi = App.Path
i = ShellExecute(Me.hWnd, vbNullString, "mailto:tesisoftware@usa.net", vbNullString, pathi, SW_SHOWNORMAL)

End Sub

Private Sub mnuLopeta_Click()

Unload frmNapit

End Sub


Private Sub mnuOhje_Click()

cmdTakaisin.Enabled = True
cmdPyoritys.Enabled = False
cmdOsta.Enabled = False
cmdRatkaise.Enabled = False
frmKirjaimet.Hide
frmTehtava.Hide
frmMain.tmrAika.Enabled = False
frmKirjaimet.tmrAika.Enabled = False
frmKirjaimet.tmrAika2.Enabled = False
frmKirjautuminen.mhCAPS.Value = False


End Sub


Private Sub mnuOhjeSisallys_Click()

Dim helppi As String
Dim X

helppi = App.Path + "\Op.hlp"
helppi = "winhlp32 " + helppi
X = Shell(helppi, 1)

End Sub

Private Sub mnuPeli_Click()

cmdTakaisin.Enabled = True
cmdPyoritys.Enabled = False
cmdOsta.Enabled = False
cmdRatkaise.Enabled = False
frmKirjaimet.Hide
frmTehtava.Hide
frmMain.tmrAika.Enabled = False
frmKirjaimet.tmrAika.Enabled = False
frmKirjaimet.tmrAika2.Enabled = False
frmKirjautuminen.mhCAPS.Value = False

End Sub


Private Sub mnuRekisterointi_Click()

frmRekisterointi.Mista = 2
frmRekisterointi.Show
frmNapit.Enabled = False
frmMain.Enabled = False

End Sub

Private Sub mnuTietoja_Click()

Load Tietojaikkuna
Tietojaikkuna.Visible = True
Tietojaikkuna.ZOrder 0
Tietojaikkuna.SetFocus

End Sub


Private Sub mnuUusipeli_Click()

Dim msg As Integer

msg = MsgBox("Haluatko varmasti uuden pelin?", 4, "Onnenpyörä")

If msg = 6 Then
    Uusipeli
End If

End Sub


