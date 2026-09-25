VERSION 5.00
Begin VB.Form frmBPalkinnot 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Bonuspalkinnot - Onnenpyörä"
   ClientHeight    =   4305
   ClientLeft      =   2265
   ClientTop       =   2160
   ClientWidth     =   6975
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4305
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK2 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Height          =   495
      Left            =   5520
      TabIndex        =   14
      Top             =   3720
      Width           =   1335
      Visible         =   0   'False
   End
   Begin VB.PictureBox pboxKuva 
      Height          =   2160
      Index           =   2
      Left            =   4680
      ScaleHeight     =   2100
      ScaleWidth      =   2100
      TabIndex        =   12
      Top             =   480
      Width           =   2160
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   5520
      TabIndex        =   11
      Top             =   3720
      Width           =   1335
   End
   Begin VB.OptionButton optVariValinta 
      Caption         =   "Vihreä"
      Height          =   255
      Index           =   2
      Left            =   3000
      TabIndex        =   10
      Top             =   3960
      Width           =   1095
   End
   Begin VB.OptionButton optVariValinta 
      Caption         =   "Keltainen"
      Height          =   255
      Index           =   1
      Left            =   3000
      TabIndex        =   9
      Top             =   3720
      Width           =   1095
   End
   Begin VB.OptionButton optVariValinta 
      Caption         =   "Punainen"
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   8
      Top             =   3480
      Width           =   1095
   End
   Begin VB.PictureBox pboxKuva 
      Height          =   2160
      Index           =   1
      Left            =   2400
      ScaleHeight     =   2100
      ScaleWidth      =   2100
      TabIndex        =   1
      Top             =   480
      Width           =   2160
   End
   Begin VB.PictureBox pboxKuva 
      Height          =   2160
      Index           =   0
      Left            =   120
      ScaleHeight     =   2100
      ScaleWidth      =   2100
      TabIndex        =   0
      Top             =   480
      Width           =   2160
   End
   Begin VB.Label lblValitse 
      Caption         =   "Valitse palkinnon väri"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Label lblPalkintoTeksti 
      Height          =   615
      Index           =   2
      Left            =   4680
      TabIndex        =   7
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label lblPalkintoTeksti 
      Height          =   615
      Index           =   1
      Left            =   2400
      TabIndex        =   6
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label lblPalkintoTeksti 
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label lblPalkinnot 
      Caption         =   "Palkinto 3"
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   4
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblPalkinnot 
      Caption         =   "Palkinto 2"
      Height          =   255
      Index           =   1
      Left            =   2520
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lblPalkinnot 
      Caption         =   "Palkinto 1"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmBPalkinnot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Palkintosarja As Integer
Public Arvo1 As Long
Public Arvo2 As Long
Public Arvo3 As Long

Public Sub Bonuspalkinnotruutuun()

If frmLoppupisteet.SuperBonus <> 1 Then
    frmBPalkinnot.Caption = "Bonuspalkinnot - Onnenpyörä"
    Randomize
    Palkintosarja = Int((4) * Rnd + 1)
    frmMain.BPalkinnonnumero1 = 29
    frmMain.BPalkinnonnumero2 = 29
    frmMain.BPalkinnonnumero3 = 29
    
    Select Case Palkintosarja

    Case 1
        lblPalkintoTeksti(0).Caption = "Yhden hengen yhden viikon matka Tunisian Sousseen, arvo 2500,-"
        lblPalkintoTeksti(1).Caption = "Yhden hengen yhden viikon matka Lontooseen, arvo 3000,-"
        lblPalkintoTeksti(2).Caption = "Matkapuhelin, arvo 3250,-"
        pboxKuva(0).Picture = frmPalkinto.pcPalkinnot.GraphicCell(29)
        pboxKuva(1).Picture = frmPalkinto.pcPalkinnot.GraphicCell(29)
        pboxKuva(2).Picture = frmPalkinto.pcPalkinnot.GraphicCell(9)
        Arvo1 = 2500
        Arvo2 = 3000
        Arvo3 = 3250
        frmMain.BPalkinnonnumero3 = 9

    Case 2
        lblPalkintoTeksti(0).Caption = "Kahden hengen yhden viikon matka Bulgariaan Sunny Beachille, arvo 5000,-"
        lblPalkintoTeksti(1).Caption = "Kahden hengen yhden viikon matka Kreetan Hanialle, arvo 6000,-"
        lblPalkintoTeksti(2).Caption = "Kahden hengen yhden viikon matka Lontooseen, arvo 6000,-"
        pboxKuva(0).Picture = frmPalkinto.pcPalkinnot.GraphicCell(29)
        pboxKuva(1).Picture = frmPalkinto.pcPalkinnot.GraphicCell(29)
        pboxKuva(2).Picture = frmPalkinto.pcPalkinnot.GraphicCell(29)
        Arvo1 = 5000
        Arvo2 = 6000
        Arvo3 = 6000

    Case 3
        lblPalkintoTeksti(0).Caption = "Yhden hengen yhden viikon matka Turkin Alanyaan, arvo 2500,-"
        lblPalkintoTeksti(1).Caption = "Matkapuhelin, arvo 3250,-"
        lblPalkintoTeksti(2).Caption = "Yhden hengen yhden viikon matka Madeiralle, arvo 4000,-"
        pboxKuva(0).Picture = frmPalkinto.pcPalkinnot.GraphicCell(29)
        pboxKuva(1).Picture = frmPalkinto.pcPalkinnot.GraphicCell(9)
        pboxKuva(2).Picture = frmPalkinto.pcPalkinnot.GraphicCell(29)
        frmMain.BPalkinnonnumero2 = 9
        Arvo1 = 2500
        Arvo2 = 3250
        Arvo3 = 4000

    Case 4
        lblPalkintoTeksti(0).Caption = "Kahden hengen yhden viikon matka Bulgariaan Sunny Beachille, arvo 5000,-"
        lblPalkintoTeksti(1).Caption = "Kahden hengen yhden viikon matka Turkin Alanyaan, arvo 5000,-"
        lblPalkintoTeksti(2).Caption = "Kahden hengen yhden viikon matka Kreetan Hanialle, arvo 6000,-"
        pboxKuva(0).Picture = frmPalkinto.pcPalkinnot.GraphicCell(29)
        pboxKuva(1).Picture = frmPalkinto.pcPalkinnot.GraphicCell(29)
        pboxKuva(2).Picture = frmPalkinto.pcPalkinnot.GraphicCell(29)
        Arvo1 = 0
        Arvo2 = 0
        Arvo3 = 0

    End Select

Else
    frmBPalkinnot.Caption = "SuperBonuspalkinnot - Onnenpyörä"
    Palkintosarja = Int((4) * Rnd + 1)
    frmMain.BPalkinnonnumero1 = 29
    frmMain.BPalkinnonnumero2 = 29
    frmMain.BPalkinnonnumero3 = 29
    
    Select Case Palkintosarja

    Case 1
        lblPalkintoTeksti(0).Caption = "Multimedia tietokone, arvo 9990,-"
        lblPalkintoTeksti(1).Caption = "Neljän hengen yhden viikon matka Fuerteventuraan, arvo 10000,-"
        lblPalkintoTeksti(2).Caption = "Neljän hengen yhden viikon matka Kyproksen Agia Napaan, arvo 11000,-"
        pboxKuva(0).Picture = frmPalkinto.pcPalkinnot.GraphicCell(30)
        pboxKuva(1).Picture = frmPalkinto.pcPalkinnot.GraphicCell(29)
        pboxKuva(2).Picture = frmPalkinto.pcPalkinnot.GraphicCell(29)
        Arvo1 = 9990
        Arvo2 = 10000
        Arvo3 = 11000
        frmMain.BPalkinnonnumero1 = 30

    Case 2
        lblPalkintoTeksti(0).Caption = "Kahden hengen kahden viikon matka Lanzaroten Playa Blancaan, arvo 8000,-"
        lblPalkintoTeksti(1).Caption = "Kahden hengen kahden viikon matka Gran Canarian Bahia Feliziin, arvo 8500,-"
        lblPalkintoTeksti(2).Caption = "Kahden hengen kahden viikon matka Madeiralle, arvo 9000,-"
        pboxKuva(0).Picture = frmPalkinto.pcPalkinnot.GraphicCell(29)
        pboxKuva(1).Picture = frmPalkinto.pcPalkinnot.GraphicCell(29)
        pboxKuva(2).Picture = frmPalkinto.pcPalkinnot.GraphicCell(29)
        Arvo1 = 8000
        Arvo2 = 8500
        Arvo3 = 9000

    Case 3
        lblPalkintoTeksti(0).Caption = "Yhden hengen neljän viikon matka Lanzaroten Playa Blancaan, arvo 9000,-"
        lblPalkintoTeksti(1).Caption = "Yhden hengen neljän viikon matka Gran Canarian Bahia Feliziin, arvo 9500,-"
        lblPalkintoTeksti(2).Caption = "Yhden hengen neljän viikon matka Madeiralle, arvo 10000,-"
        pboxKuva(0).Picture = frmPalkinto.pcPalkinnot.GraphicCell(29)
        pboxKuva(1).Picture = frmPalkinto.pcPalkinnot.GraphicCell(29)
        pboxKuva(2).Picture = frmPalkinto.pcPalkinnot.GraphicCell(29)
        Arvo1 = 0
        Arvo2 = 0
        Arvo3 = 0

    Case 4
        lblPalkintoTeksti(0).Caption = "Multimedia tietokone, arvo 9990,-"
        lblPalkintoTeksti(1).Caption = "Kahden hengen kahden viikon matka Fuerteventuraan, arvo 10500,-"
        lblPalkintoTeksti(2).Caption = "Kahden hengen kahden viikon matka Kyproksen Agia Napaan, arvo 11500,-"
        pboxKuva(0).Picture = frmPalkinto.pcPalkinnot.GraphicCell(30)
        pboxKuva(1).Picture = frmPalkinto.pcPalkinnot.GraphicCell(29)
        pboxKuva(2).Picture = frmPalkinto.pcPalkinnot.GraphicCell(29)
        Arvo1 = 0
        Arvo2 = 0
        Arvo3 = 0
        frmMain.BPalkinnonnumero1 = 30

    End Select

End If

End Sub
Private Sub cmdOK_Click()

frmNapit.cmdPyoritys.Enabled = False
frmNapit.cmdOsta.Enabled = False
frmNapit.cmdRatkaise.Enabled = False
frmKirjaimet.Enabled = True

If optVariValinta(0).Value = True Then
    frmMain.PalkintoVari = 1
End If

If optVariValinta(1).Value = True Then
    frmMain.PalkintoVari = 2
End If

If optVariValinta(2).Value = True Then
    frmMain.PalkintoVari = 3
End If

frmMain.AloitaBonus
frmMain.BPalkinnonteksti1 = lblPalkintoTeksti(0).Caption
frmMain.BPalkinnonteksti2 = lblPalkintoTeksti(1).Caption
frmMain.BPalkinnonteksti3 = lblPalkintoTeksti(2).Caption
Unload frmBPalkinnot

End Sub
Private Sub cmdOK2_Click()

frmKirjautuminen.mhCAPS = False
Dim msg

msg = MsgBox("Haluatko pelata uudelleen?", vbYesNo, "Onnenpyörä")

If msg = 6 Then
    frmNapit.Uusipeli

Else
    End

End If

End Sub

Private Sub optVariValinta_Click(Index As Integer)
    
cmdOK.Enabled = True

End Sub


