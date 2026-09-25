VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form frmPalkinto 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Palkinto - Onnenpyörä"
   ClientHeight    =   3555
   ClientLeft      =   2910
   ClientTop       =   2625
   ClientWidth     =   6195
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3555
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK2 
      Caption         =   "&OK"
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      Top             =   2760
      Width           =   1575
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   615
      Left            =   2160
      TabIndex        =   2
      Top             =   2760
      Width           =   1575
   End
   Begin VB.PictureBox pboxKuva 
      Height          =   2160
      Left            =   3720
      ScaleHeight     =   2100
      ScaleWidth      =   2100
      TabIndex        =   0
      Top             =   360
      Width           =   2160
   End
   Begin PicClip.PictureClip pcPalkinnot 
      Left            =   0
      Top             =   5000
      _ExtentX        =   11113
      _ExtentY        =   40746
      _Version        =   327680
      Rows            =   11
      Cols            =   3
      Picture         =   "frmPalkinto.frx":0000
   End
   Begin VB.Label Label1 
      Caption         =   "Pistetilanne:"
      Height          =   255
      Left            =   1320
      TabIndex        =   10
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblPelaaja1 
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   2295
   End
   Begin VB.Label lblPisteet1 
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   1680
      Width           =   675
   End
   Begin VB.Label lblPisteet2 
      Height          =   255
      Left            =   2760
      TabIndex        =   7
      Top             =   2040
      Width           =   675
   End
   Begin VB.Label lblPisteet3 
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   2400
      Width           =   675
   End
   Begin VB.Label lblPelaaja2 
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   2295
   End
   Begin VB.Label lblPelaaja3 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   2295
   End
   Begin VB.Label lblTeksti 
      Height          =   855
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "frmPalkinto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Palkintonumero As Integer
Public Palkinto As String

Public Sub Palkinnot()

Randomize
Palkintonumero = Int((4) * Rnd + 1)

Select Case Palkintonumero

Case 1
    
    Select Case frmMain.Kierros
    
    Case 2
        Palkinto = "vedenkeittimen, jonka arvo on 395mk."
        pboxKuva.Picture = pcPalkinnot.GraphicCell(27)
        frmMain.Arvo = 395
    
    Case 3
        Palkinto = "20h risteilyn, jonka arvo on 840mk."
        pboxKuva.Picture = pcPalkinnot.GraphicCell(18)
        frmMain.Arvo = 840
    
    Case 4
        Palkinto = "rannekellon, jonka arvo on 745mk."
        pboxKuva.Picture = pcPalkinnot.GraphicCell(22)
        frmMain.Arvo = 745
    
    Case 5
        Palkinto = "matkapuhelimen, jonka arvo on 1695mk."
        pboxKuva.Picture = pcPalkinnot.GraphicCell(10)
        frmMain.Arvo = 1695
    
    Case 6
        Palkinto = "minihifisarjan, jonka arvo on 1795mk."
        pboxKuva.Picture = pcPalkinnot.GraphicCell(2)
        frmMain.Arvo = 1795
    End Select

Case 2
    
    Select Case frmMain.Kierros
    
    Case 2
        Palkinto = "silitysraudan, jonka arvo on 395mk."
        pboxKuva.Picture = pcPalkinnot.GraphicCell(6)
        frmMain.Arvo = 395
    
    Case 3
        Palkinto = "matkapuhelimen, jonka arvo on 695mk."
        pboxKuva.Picture = pcPalkinnot.GraphicCell(13)
        frmMain.Arvo = 695
    
    Case 4
        Palkinto = "kameran, jonka arvo on 795mk."
        pboxKuva.Picture = pcPalkinnot.GraphicCell(14)
        frmMain.Arvo = 795
    
    Case 5
        Palkinto = "matkatelevision, jonka arvo on 1395mk."
        pboxKuva.Picture = pcPalkinnot.GraphicCell(23)
        frmMain.Arvo = 1395
    
    Case 6
        Palkinto = "hifi-videonauhurin, jonka arvo on 2195mk."
        pboxKuva.Picture = pcPalkinnot.GraphicCell(17)
        frmMain.Arvo = 2195
    
    End Select

Case 3
    
    Select Case frmMain.Kierros
    
    Case 2
        Palkinto = "sähköhammasharjan, jonka arvo on 495mk."
        pboxKuva.Picture = pcPalkinnot.GraphicCell(4)
        frmMain.Arvo = 495
    
    Case 3
        Palkinto = "CD-radio nauhurin, jonka arvo on 695mk."
        pboxKuva.Picture = pcPalkinnot.GraphicCell(0)
        frmMain.Arvo = 695
    
    Case 4
        Palkinto = "pölynimurin, jonka arvo on 895mk."
        pboxKuva.Picture = pcPalkinnot.GraphicCell(16)
        frmMain.Arvo = 895
    
    Case 5
        Palkinto = "kuntopyörän, jonka arvo on 1995mk."
        pboxKuva.Picture = pcPalkinnot.GraphicCell(8)
        frmMain.Arvo = 1995
    
    Case 6
        Palkinto = "matkapuhelimen, jonka arvo on 2495mk."
        pboxKuva.Picture = pcPalkinnot.GraphicCell(12)
        frmMain.Arvo = 2495
    
    End Select

Case 4
    
    Select Case frmMain.Kierros
    
    Case 2
        Palkinto = "kannettavan CD-soittimen, jonka arvo on 595mk."
        pboxKuva.Picture = pcPalkinnot.GraphicCell(19)
        frmMain.Arvo = 595
    
    Case 3
        Palkinto = "johdottoman puhelimen, jonka arvo on 695mk."
        pboxKuva.Picture = pcPalkinnot.GraphicCell(20)
        frmMain.Arvo = 695
    
    Case 4
        Palkinto = "videonauhurin, jonka arvo on 995mk."
        pboxKuva.Picture = pcPalkinnot.GraphicCell(28)
        frmMain.Arvo = 995
    
    Case 5
        Palkinto = "kasvosolariumin, jonka arvo on 1495mk."
        pboxKuva.Picture = pcPalkinnot.GraphicCell(11)
        frmMain.Arvo = 1495
    
    Case 6
        Palkinto = "20 " + Chr$(34) + " väritelevision, jonka arvo on 2695mk."
        pboxKuva.Picture = pcPalkinnot.GraphicCell(5)
        frmMain.Arvo = 2695
    
    End Select

End Select

End Sub

Public Sub YPalkinnot()

Randomize
Palkintonumero = Int((4) * Rnd + 1)

Select Case Palkintonumero

Case 1
    Palkinto = "kahvinkeittimen, jonka arvo on 295mk."
    pboxKuva.Picture = pcPalkinnot.GraphicCell(1)
    frmMain.Arvo = 295

Case 2
    Palkinto = "sauvasekoittimen, jonka arvo on 295mk."
    pboxKuva.Picture = pcPalkinnot.GraphicCell(3)
    frmMain.Arvo = 295

Case 3
    Palkinto = "kelloradion, jonka arvo on 395mk."
    pboxKuva.Picture = pcPalkinnot.GraphicCell(15)
    frmMain.Arvo = 395

Case 4
    Palkinto = "hiustenkuivaajan, jonka arvo on 295mk."
    pboxKuva.Picture = pcPalkinnot.GraphicCell(24)
    frmMain.Arvo = 295

End Select

End Sub

Public Sub SYPalkinnot()

Randomize
Palkintonumero = Int((4) * Rnd + 1)

Select Case Palkintonumero

Case 1
    Palkinto = "parranajokoneen, jonka arvo on 495mk."
    pboxKuva.Picture = pcPalkinnot.GraphicCell(7)
    frmMain.Arvo = 495

Case 2
    Palkinto = "leiväpaahtimen, jonka arvo on 595mk."
    pboxKuva.Picture = pcPalkinnot.GraphicCell(25)
    frmMain.Arvo = 595

Case 3
    Palkinto = "langattomat kuulokkeet, jonka arvo on 595mk."
    pboxKuva.Picture = pcPalkinnot.GraphicCell(26)
    frmMain.Arvo = 595

Case 4
    Palkinto = "yleiskoneen, jonka arvo on 595mk."
    pboxKuva.Picture = pcPalkinnot.GraphicCell(21)
    frmMain.Arvo = 595

End Select

End Sub


Public Sub Teksti(i As Integer)

Dim Tex As String

Select Case i

Case 1
    Palkinnot
    Tex = ". Olet voittanut "

Case 2
    YPalkinnot
    Tex = ". Yllätys palkintona voitat "

Case 3
    SYPalkinnot
    Tex = ". Superyllätys palkintona voitat "

Case 4
    Palkinnot
    Tex = ". Tehtävä palkintona voitat "

End Select

Select Case frmMain.Pelaaja

Case 1
    lblTeksti.Caption = "Onneksi olkoon, " + frmMain.lblP1.Caption + Tex + Palkinto
    frmMain.RahaMaaraP1 = frmMain.RahaMaaraP1 + frmMain.Arvo

Case 2
    lblTeksti.Caption = "Onneksi olkoon, " + frmMain.lblP2.Caption + Tex + Palkinto
    frmMain.RahaMaaraP2 = frmMain.RahaMaaraP2 + frmMain.Arvo

Case 3
    lblTeksti.Caption = "Onneksi olkoon, " + frmMain.lblP3.Caption + Tex + Palkinto
    frmMain.RahaMaaraP3 = frmMain.RahaMaaraP3 + frmMain.Arvo

End Select

End Sub

Private Sub cmdOK_Click()

If frmMain.Palkinnonaika <> 1 Then
    If frmMain.Kierros = 6 Then
        frmLoppupisteet.Nimet
        frmLoppupisteet.Show
        frmLoppupisteet.SetFocus
        frmPalkinto.Hide
        Exit Sub
    End If
    
    frmMain.Kelaus = 1
    frmMain.TehtavanAloitus
    frmMain.tmrAika.Enabled = True
    frmNapit.Enabled = True
    frmNapit.cmdTakaisin_Click
    frmTehtava.Visible = True
    frmKirjaimet.Visible = True
    frmNapit.SetFocus
    Unload frmPalkinto

Else
    frmMain.Palkinnonaika = 0
    frmMain.tmrAika.Enabled = True
    Unload frmPalkinto
    frmTehtava.Visible = True
    frmKirjaimet.Visible = True
    frmNapit.Enabled = True
    frmNapit.cmdTakaisin_Click
    frmNapit.SetFocus

End If

End Sub


Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    cmdOK_Click
End If

End Sub


Private Sub cmdOK2_Click()

frmMain.PisteetP1 = 0
frmMain.PisteetP2 = 0
frmMain.PisteetP3 = 0
frmMain.Pisteet
Teksti (4)
cmdOK.Visible = True
cmdOK2.Visible = False

End Sub

Private Sub cmdOK2_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    cmdOK2_Click
End If

End Sub


Private Sub Form_Load()

frmNapit.Enabled = False
frmMain.Enabled = False
frmKirjaimet.Enabled = False
frmTehtava.Enabled = False

If frmMain.Palkinto = 3 Then
    cmdOK.Visible = False
    cmdOK2.Visible = True
Else
    cmdOK.Visible = True
    cmdOK2.Visible = False
End If

Teksti (frmMain.Palkinto)

Label1.Visible = False

If frmMain.Palkinnonaika <> 1 Then

Select Case frmKirjautuminen.Pluku

Case 1
    lblPelaaja1.Caption = frmMain.lblP1.Caption
    lblPisteet1.Caption = frmMain.KokoPisteetP1

Case 2
    lblPelaaja1.Caption = frmMain.lblP1.Caption
    lblPisteet1.Caption = frmMain.KokoPisteetP1
    lblPelaaja2.Caption = frmMain.lblP2.Caption
    lblPisteet2.Caption = frmMain.KokoPisteetP2

Case 3
    lblPelaaja1.Caption = frmMain.lblP1.Caption
    lblPisteet1.Caption = frmMain.KokoPisteetP1
    lblPelaaja2.Caption = frmMain.lblP2.Caption
    lblPisteet2.Caption = frmMain.KokoPisteetP2
    lblPelaaja3.Caption = frmMain.lblP3.Caption
    lblPisteet3.Caption = frmMain.KokoPisteetP3

End Select

Label1.Visible = True

End If

End Sub
