VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C00000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Onnenpyörä"
   ClientHeight    =   5040
   ClientLeft      =   1515
   ClientTop       =   5370
   ClientWidth     =   5655
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5040
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox picPyora 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   240
      ScaleHeight     =   2535
      ScaleWidth      =   5085
      TabIndex        =   3
      Top             =   2520
      Width           =   5085
      Begin VB.Timer tmrAika 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   840
         Top             =   480
      End
   End
   Begin PicClip.PictureClip pcKuva 
      Left            =   600
      Top             =   5160
      _ExtentX        =   17992
      _ExtentY        =   13494
      _Version        =   327680
      Rows            =   3
      Cols            =   2
      Picture         =   "frmMain.frx":0E42
   End
   Begin VB.Label lblSektori 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1920
      TabIndex        =   7
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   3375
      Left            =   0
      Top             =   1690
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   135
      Left            =   -240
      Top             =   1560
      Width           =   5895
   End
   Begin VB.Shape Valo1 
      BackColor       =   &H00008080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Height          =   375
      Left            =   120
      Shape           =   3  'Circle
      Top             =   240
      Width           =   375
   End
   Begin VB.Label lblP3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   4440
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblP2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   2520
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lblP1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   3375
      Left            =   5280
      Top             =   1680
      Width           =   375
   End
   Begin VB.Label lblPisteet3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblPisteet2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label lblPisteet1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
   Begin VB.Shape Valo2 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Height          =   375
      Left            =   2040
      Shape           =   3  'Circle
      Top             =   240
      Width           =   375
   End
   Begin VB.Shape Valo3 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   5
      Height          =   375
      Left            =   3960
      Shape           =   3  'Circle
      Top             =   240
      Width           =   375
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FF0000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FF0000&
      Height          =   820
      Left            =   255
      Top             =   1695
      Width           =   5040
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Kuva As Integer
Public Sektori As Integer
Public YllatysSaatu As Integer
Public SYllatysSaatu As Integer
Public YllatyksenVuoro As Integer
Public Sektoriteksti As String
Public Kierros As Integer
Public KokoPisteetP1 As Long
Public KokoPisteetP2 As Long
Public KokoPisteetP3 As Long
Public PisteetP1 As Long
Public PisteetP2 As Long
Public PisteetP3 As Long
Public RahaMaaraP1 As Long
Public RahaMaaraP2 As Long
Public RahaMaaraP3 As Long
Public Arvo As Long
Public Tehtava1 As Integer
Public Tehtava2 As Integer
Public Tehtava3 As Integer
Public Tehtava4 As Integer
Public Tehtava5 As Integer
Public UusiTehtava As Integer
Public Pelaaja As Integer
Public SektoriPisteet As Integer
Public Vokaaliostettu As Integer
Public Aika As Integer
Public Kelaus As Integer
Public numero As Integer
Public maxnumero As Integer
Public maxtnumero As Integer
Public Palkinto As Integer
Public Palkinnonaika As Integer
Public SPalkinnonsaaja As Integer
Public PalkintoVari As Integer
Public BPalkinnonnumero1 As Integer
Public BPalkinnonteksti1 As String
Public BPalkinnonnumero2 As Integer
Public BPalkinnonteksti2 As String
Public BPalkinnonnumero3 As Integer
Public BPalkinnonteksti3 As String
Public TodellaLoppu As Integer

Public Sub AloitaBonus()

Randomize
TehtavaArpa = Int((maxnumero) * Rnd + 1)
UusiTehtava = TehtavaArpa
frmKirjaimet.BonusAloitus (UusiTehtava)

End Sub

Public Sub Aloitusvuoro()

Select Case Kierros

Case 1
        Pelaaja = 1
Case 2
    
    Select Case frmKirjautuminen.Pluku
    
    Case 1
        Pelaaja = 1
    
    Case 2
        Pelaaja = 2
    
    Case 3
        Pelaaja = 2
    
    End Select

Case 3
    
    Select Case frmKirjautuminen.Pluku
    
    Case 1
        Pelaaja = 1
    
    Case 2
        Pelaaja = 1
    
    Case 3
        Pelaaja = 3
    
    End Select

Case 4
    
    Select Case frmKirjautuminen.Pluku
    
    Case 1
        Pelaaja = 1
    
    Case 2
        Pelaaja = 2
    
    Case 3
        Pelaaja = 1
    
    End Select

Case 5
    
    Select Case frmKirjautuminen.Pluku
    
    Case 1
        Pelaaja = 1
    
    Case 2
        Pelaaja = 1
    
    Case 3
        Pelaaja = 2
    
    End Select

End Select

Select Case Pelaaja

Case 1
    Valo1.BackColor = &HFFFF&
    Valo2.BackColor = &H80&
    Valo3.BackColor = &H800080

Case 2
    Valo1.BackColor = &H8080&
    Valo2.BackColor = &HFF&
    Valo3.BackColor = &H800080

Case 3
    Valo1.BackColor = &H8080&
    Valo2.BackColor = &H80&
    Valo3.BackColor = &HFF00FF

End Select

End Sub
Public Sub BonusTehtava()

Dim BonusTehtavaArpa As Integer

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

For i = 0 To 28
    frmKirjaimet.lblValitut(i).Visible = True
Next

frmLoppupisteet.Hide
frmBPalkinnot.Bonuspalkinnotruutuun
frmBPalkinnot.Show

End Sub

Public Sub MaxTehtnumero()

Dim tiedosto As String
Dim texti As Integer

tiedosto = App.Path + "\teht.txt"
maxtnumero = 0

Open tiedosto For Input As #1
    
    Do
        
        Line Input #1, rivi
            texti = Left(rivi, 3)
    
    Loop Until EOF(1)

Close #1

maxtnumero = Int(texti)

End Sub

Public Sub MaxBonusTehtnumero()

Dim tiedosto As String
Dim texti As Integer

tiedosto = App.Path + "\bteht.txt"
maxnumero = 0

Open tiedosto For Input As #1
    
    Do
        
        Line Input #1, rivi
            texti = Left(rivi, 3)
    
    Loop Until EOF(1)

Close #1

maxnumero = Int(texti)

End Sub


Private Sub Arvonta()

Dim Arpa As Integer

frmMain.YllatyksenVuoro = 0
Arpa = Int((24) * Rnd + 1)
'Arpa = 7

Select Case Arpa

Case 1, 2
    Sektori = 1 'rosvo

Case 3, 4, 5, 6
    Sektori = 2 'ohi

Case 7
    
    Select Case Kierros
    
    Case 2
        If YllatysSaatu = 0 Then
                Sektori = 3 'yllätys
                YllatyksenVuoro = 1
        
        Else
                Sektori = 5 '450
        
        End If
    
    Case 3
        If SYllatysSaatu = 0 Then
                Sektori = 4 'superyllätys
                YllatyksenVuoro = 1
        
        Else
                Sektori = 5 '450
        
        End If
    
    Case Else
        Sektori = 5 '450
    
    End Select

Case Else
    Sektori = Arpa - 2

End Select

Muuta_Sektorit

End Sub

Public Sub KonsonantitNakyviin()

For i = 0 To 19
    frmKirjaimet.cmdKonso(i).Enabled = True

Next

frmNapit.Enabled = False

End Sub

Public Sub Pisteet()

lblPisteet1.Caption = PisteetP1
lblPisteet2.Caption = PisteetP2
lblPisteet3.Caption = PisteetP3

End Sub

Public Sub Ratkaistu()

Dim msg As Integer

Kierros = Kierros + 1

Select Case Pelaaja

Case 1
    If PisteetP1 = 0 Then
        PisteetP1 = PisteetP1 + 300
    
    End If
    
    KokoPisteetP1 = KokoPisteetP1 + PisteetP1

Case 2
    
    If PisteetP2 = 0 Then
        PisteetP2 = PisteetP2 + 300
    
    End If
    
    KokoPisteetP2 = KokoPisteetP2 + PisteetP2

Case 3
    If PisteetP3 = 0 Then
        PisteetP3 = PisteetP3 + 300
    
    End If
    
    KokoPisteetP3 = KokoPisteetP3 + PisteetP3

End Select

lblPisteet1.Caption = KokoPisteetP1
lblPisteet2.Caption = KokoPisteetP2
lblPisteet3.Caption = KokoPisteetP3

If Kierros <> 6 Then
    frmMain.tmrAika.Enabled = False
    msg = MsgBox("Onneksi olkoon ratkaisit tehtävän.", 0, "Onnenpyörä")
    
    If SYllatysSaatu = 1 And Pelaaja = SPalkinnonsaaja Then
        Palkinto = 3
        frmPalkinto.Show
    Else
    
    Palkinto = 1
    frmPalkinto.Show
    PisteetP1 = 0
    PisteetP2 = 0
    PisteetP3 = 0
    Pisteet
    
    End If

Else
    frmMain.tmrAika.Enabled = False
    msg = MsgBox("Onneksi olkoon ratkaisit tehtävän.", 0, "Onnenpyörä")
    Palkinto = 1
    frmPalkinto.Show
End If

End Sub

Public Sub SuperYllatys()

Dim msg As Integer

frmTehtava.Visible = False
frmKirjaimet.Visible = False
msg = MsgBox("Jos onnistut ratkaisemaan tehtävän, saat tehtäväpalkinnon lisäksi Superyllätyspalkinnon.", 0, "Onnenpyörä")
frmTehtava.Visible = True
frmKirjaimet.Visible = True
SPalkinnonsaaja = Pelaaja
SYllatysSaatu = 1
frmNapit.cmdOsta.Enabled = True
frmNapit.cmdRatkaise.Enabled = True
frmNapit.cmdPyoritys.Enabled = True
frmMain.Aika = 7
frmMain.tmrAika.Enabled = True

End Sub


Public Sub VokaalitNakyviin()

For i = 0 To 8
    frmKirjaimet.cmdVokaali(i).Enabled = True
Next

frmNapit.Enabled = False
frmKirjaimet.Enabled = True
frmKirjaimet.SetFocus

End Sub

Public Sub KonsonantitPiiloon()

For i = 0 To 19
    frmKirjaimet.cmdKonso(i).Enabled = False

Next

End Sub


Public Sub VokaalitPiiloon()

For i = 0 To 8
    frmKirjaimet.cmdVokaali(i).Enabled = False

Next

End Sub

Private Sub Muuta_Sektorit()

Select Case Sektori

Case 1
    Sektoriteksti = "ROSVO"

Case 2
    Sektoriteksti = "OHI"

Case 3
    Sektoriteksti = "YLLÄTYS"

Case 4
    Sektoriteksti = "SUPER YLLÄTYS"

Case 5
    Sektoriteksti = "450"

Case 6
    Sektoriteksti = "1600"

Case 7
    Sektoriteksti = "1350"

Case 8
    Sektoriteksti = "1300"

Case 9
    Sektoriteksti = "950"

Case 10
    Sektoriteksti = "900"

Case 11
    Sektoriteksti = "850"

Case 12
    Sektoriteksti = "800"

Case 13
    Sektoriteksti = "750"

Case 14
    Sektoriteksti = "700"

Case 15
    Sektoriteksti = "650"

Case 16
    Sektoriteksti = "600"

Case 17
    Sektoriteksti = "500"

Case 18
    Sektoriteksti = "400"

Case 19
    Sektoriteksti = "350"

Case 20
    Sektoriteksti = "300"

Case 21
    Sektoriteksti = "250"

Case 22
    Sektoriteksti = "100"

End Select

End Sub


Public Sub TehtavanAloitus()

Dim TehtavaArpa As Integer

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

For i = 0 To 28
    frmKirjaimet.lblValitut(i).Visible = True
Next
Aloitusvuoro
Arvonta:
Randomize
TehtavaArpa = Int((maxtnumero) * Rnd + 1)

If TehtavaArpa <> Tehtava1 And TehtavaArpa <> Tehtava2 And TehtavaArpa <> Tehtava3 And TehtavaArpa <> Tehtava4 Then
    UusiTehtava = TehtavaArpa

Else
    GoTo Arvonta

End If

Select Case Kierros

Case 1
    Tehtava1 = UusiTehtava

Case 2
    Tehtava2 = UusiTehtava

Case 3
    Tehtava3 = UusiTehtava

Case 4
    Tehtava4 = UusiTehtava

End Select

frmKirjaimet.Aloitus (UusiTehtava)
frmMain.tmrAika.Enabled = True

If frmMain.Aika > 5 Then
    frmMain.Aika = 20
Else
    frmMain.Aika = 7
End If

End Sub

Public Sub Pyoraytys()

For J = 1 To 15
    
    For i = 0 To 5
        picPyora.Picture = pcKuva.GraphicCell(i)
        lblSektori.Refresh
        
        For K = 1 To 150000
        Next
    
    Next

Arvonta
lblSektori = Sektoriteksti

Next

Sektorit_Pisteiksi

If Sektori <> 1 And Sektori <> 2 And Sektori <> 3 And Sektori <> 4 Then
    frmKirjaimet.Enabled = True
    frmKirjaimet.SetFocus
    frmMain.tmrAika.Enabled = True
    frmMain.Aika = 7
    KonsonantitNakyviin
Else

    Select Case Sektori
    
    Case 1 'rosvo
        frmTehtava.Visible = False
        frmKirjaimet.Visible = False
        msg = MsgBox("Rosvo sektori. Kaikki pisteet menevät!", 0, "Onnenpyörä")
        frmTehtava.Visible = True
        frmKirjaimet.Visible = True
        
        Select Case Pelaaja
        
        Case 1
            PisteetP1 = 0
        
        Case 2
            PisteetP2 = 0
        
        Case 3
            PisteetP3 = 0
        
        End Select
        
        Pisteet
        VaihdaVuoro
        frmNapit.cmdPyoritys.Enabled = True
        frmNapit.cmdOsta.Enabled = True
        frmNapit.cmdRatkaise.Enabled = True
    
    Case 2 'ohi
        frmTehtava.Visible = False
        frmKirjaimet.Visible = False
        msg = MsgBox("Ohi sektori. Vuoro vaihtuu.", 0, "Onnenpyörä")
        frmTehtava.Visible = True
        frmKirjaimet.Visible = True
        VaihdaVuoro
        frmNapit.cmdPyoritys.Enabled = True
        frmNapit.cmdOsta.Enabled = True
        frmNapit.cmdRatkaise.Enabled = True
    
    Case 3 'yllätys
        
        Yllatys
        YllatysSaatu = 1
    
    Case 4 'superyllätys
        
        SuperYllatys
        SYllatysSaatu = 1
    
    End Select
    
    frmMain.Aika = 7

End If

End Sub

Public Sub Sektorit_Pisteiksi()

Select Case Sektori

Case 5
    SektoriPisteet = 450

Case 6
    SektoriPisteet = 1600

Case 7
    SektoriPisteet = 1350

Case 8
    SektoriPisteet = 1300

Case 9
    SektoriPisteet = 950

Case 10
    SektoriPisteet = 900

Case 11
    SektoriPisteet = 850

Case 12
    SektoriPisteet = 800

Case 13
    SektoriPisteet = 750

Case 14
    SektoriPisteet = 700

Case 15
    SektoriPisteet = 650

Case 16
    SektoriPisteet = 600

Case 17
    SektoriPisteet = 500

Case 18
    SektoriPisteet = 400

Case 19
    SektoriPisteet = 350

Case 20
    SektoriPisteet = 300

Case 21
    SektoriPisteet = 250

Case 22
    SektoriPisteet = 100

End Select

End Sub


Public Sub Yllatys()

Dim msg As Integer

frmMain.tmrAika.Enabled = False
msg = MsgBox("Jos arvaat oikein yhden vokaalin tai konsonantin voitat yllätyspalkinnon.", 0, "Onnenpyörä")
YllatysSaatu = 1
Palkinnonaika = 1
VokaalitNakyviin
KonsonantitNakyviin

End Sub


Private Sub Form_Load()

frmKirjautuminen.Hide
frmLoppupisteet.Hide
Select Case Screen.Width / 15
Case 640
    frmMain.Top = 3000
    frmMain.Left = -100
Case 800
    frmMain.Top = 3200
    frmMain.Left = 0
Case Else
    frmMain.Top = 5000
    frmMain.Left = 1500
End Select
Load frmTehtava
frmTehtava.Show
Load frmKirjaimet
frmKirjaimet.Show
Load frmNapit
frmNapit.Show
picPyora.Picture = pcKuva.GraphicCell(5)
Kierros = 1
Kuva = 1
MaxTehtnumero
MaxBonusTehtnumero
lblP1.Caption = frmKirjautuminen.Nimi1
lblP2.Caption = frmKirjautuminen.Nimi2
lblP3.Caption = frmKirjautuminen.Nimi3
frmNapit.Enabled = True
frmMain.Enabled = False
frmKirjaimet.Enabled = False
frmTehtava.Enabled = False
frmNapit.SetFocus
PisteetP1 = 0
PisteetP2 = 0
PisteetP3 = 0
lblPisteet1 = PisteetP1
lblPisteet2 = PisteetP2
lblPisteet3 = PisteetP3
Pelaaja = 3
Vokaaliostettu = 0
Palkinnonaika = 0
Aika = 20
TehtavanAloitus
TodellaLoppu = 1
End Sub
Public Sub VaihdaVuoro()

frmMain.VokaalitPiiloon
frmMain.KonsonantitPiiloon
frmNapit.cmdOsta.Enabled = True
frmNapit.cmdPyoritys.Enabled = True
frmNapit.cmdRatkaise.Enabled = True
frmMain.tmrAika.Enabled = True

If frmMain.Aika > 5 Then
    frmMain.Aika = 20
Else
    frmMain.Aika = 7
End If

Select Case Pelaaja

Case 1
    
    Select Case frmKirjautuminen.Pluku
    
    Case 1
        Exit Sub
    
    Case 2, 3
        Pelaaja = 2
    
    End Select

Case 2
    
    Select Case frmKirjautuminen.Pluku
    
    Case 1
        Pelaaja = 1
    
    Case 2
        Pelaaja = 1
    
    Case 3
        Pelaaja = 3
    
    End Select

Case 3
    
    Select Case frmKirjautuminen.Pluku
    
    Case 1
        Pelaaja = 1
    
    Case 2
        Pelaaja = 1
    
    Case 3
        Pelaaja = 1
    
    End Select

End Select

Select Case Pelaaja

Case 1
    Valo1.BackColor = &HFFFF&
    Valo2.BackColor = &H80&
    Valo3.BackColor = &H800080

Case 2
    Valo1.BackColor = &H8080&
    Valo2.BackColor = &HFF&
    Valo3.BackColor = &H800080

Case 3
    Valo1.BackColor = &H8080&
    Valo2.BackColor = &H80&
    Valo3.BackColor = &HFF00FF

End Select

frmNapit.Enabled = True
frmKirjaimet.Enabled = False
frmNapit.SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)

If TodellaLoppu = 1 Then
    End

End If

End Sub


Private Sub tmrAika_Timer()

Dim msg As Integer

Aika = Aika - 1

If Aika < 1 Then
    frmTehtava.Visible = False
    frmKirjaimet.Visible = False
    msg = MsgBox("Aika loppui.", 0, "Onnenpyörä")
    frmTehtava.Visible = True
    frmKirjaimet.Visible = True
    VaihdaVuoro
    frmNapit.Enabled = True
    frmKirjaimet.Enabled = False
    frmNapit.SetFocus
End If

End Sub


