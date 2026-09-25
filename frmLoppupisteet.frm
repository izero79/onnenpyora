VERSION 5.00
Begin VB.Form frmLoppupisteet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Loppupisteet - Onnenpyörä"
   ClientHeight    =   2880
   ClientLeft      =   2880
   ClientTop       =   3015
   ClientWidth     =   5325
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2880
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   615
      Left            =   3240
      TabIndex        =   6
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label lblValmistaudu 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label lblPelaaja3 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   2895
   End
   Begin VB.Label lblPelaaja2 
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label lblPisteet3 
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label lblPisteet2 
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label lblPisteet1 
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblPelaaja1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmLoppupisteet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BonukseenMenija As Integer
Public SuperBonus As Integer
Public SBonus As String
Public Arvonta As Integer

Public Sub Nimet()

Dim Arvopa As Integer
Dim vteksti As String

frmMain.tmrAika.Enabled = False
frmKirjaimet.tmrAika.Enabled = False

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

frmNapit.Enabled = False
frmLoppupisteet.Enabled = True
SuperBonus = 0

If frmMain.KokoPisteetP1 > frmMain.KokoPisteetP2 And frmMain.KokoPisteetP1 > frmMain.KokoPisteetP3 Then
    
    If frmMain.KokoPisteetP2 = 0 And frmMain.KokoPisteetP3 = 0 Then
        SuperBonus = 1
    
    End If
    
    BonukseenMenija = 1

End If

If frmMain.KokoPisteetP2 > frmMain.KokoPisteetP1 And frmMain.KokoPisteetP2 > frmMain.KokoPisteetP3 Then
    
    If frmMain.KokoPisteetP1 = 0 And frmMain.KokoPisteetP3 = 0 Then
        SuperBonus = 1
    
    End If
    
    BonukseenMenija = 2

End If

If frmMain.KokoPisteetP3 > frmMain.KokoPisteetP2 And frmMain.KokoPisteetP3 > frmMain.KokoPisteetP1 Then
    
    If frmMain.KokoPisteetP2 = 0 And frmMain.KokoPisteetP1 = 0 Then
        SuperBonus = 1
    
    End If
    
    BonukseenMenija = 3

End If

Arvonta = 0

If frmMain.KokoPisteetP1 = frmMain.KokoPisteetP2 And frmMain.KokoPisteetP1 > frmMain.KokoPisteetP3 Then
    Arvonta = 3

End If

If frmMain.KokoPisteetP1 = frmMain.KokoPisteetP3 And frmMain.KokoPisteetP1 > frmMain.KokoPisteetP2 Then
    Arvonta = 4

End If

If frmMain.KokoPisteetP2 = frmMain.KokoPisteetP3 And frmMain.KokoPisteetP3 > frmMain.KokoPisteetP1 Then
    Arvonta = 5

End If

If frmMain.KokoPisteetP2 = frmMain.KokoPisteetP3 And frmMain.KokoPisteetP3 = frmMain.KokoPisteetP1 Then
    Arvonta = 6

End If

If SuperBonus = 1 Then
    SBonus = "SUPERBONUKSEEN!"

Else
    SBonus = "BONUS tehtävään!"

End If

Select Case Arvonta

Case 3
    
    msg = MsgBox("Koska kahdella pelaajalla on sama pistemäärä suoritetaan arvonta voitosta näiden kahden pelaajan välillä.", 0, "Onnenpyörä")
    Randomize
    Arvopa = Int((2) * Rnd + 1)
    
    Select Case Arvopa
    
    Case 1
        vteksti = "Voittaja on pelaaja 1 eli " + lblPelaaja1.Caption + "."
        BonukseenMenija = 1
    
    Case 2
        vteksti = "Voittaja on pelaaja 2 eli " + lblPelaaja2.Caption + "."
        BonukseenMenija = 2
    
    End Select
    
    msg = MsgBox(vteksti, 0, "Onnenpyörä")
    
Case 4
    
    msg = MsgBox("Koska kahdella pelaajalla on sama pistemäärä suoritetaan arvonta voitosta näiden kahden pelaajan välillä.", 0, "Onnenpyörä")
    Randomize
    Arvopa = Int((2) * Rnd + 1)
    
    Select Case Arvopa
    
    Case 1
        vteksti = "Voittaja on pelaaja 1 eli " + lblPelaaja1.Caption + "."
        BonukseenMenija = 1
    
    Case 2
        vteksti = "Voittaja on pelaaja 3 eli " + lblPelaaja3.Caption + "."
        BonukseenMenija = 3
    
    End Select
    
    msg = MsgBox(vteksti, 0, "Onnenpyörä")

Case 5
    
    msg = MsgBox("Koska kahdella pelaajalla on sama pistemäärä suoritetaan arvonta voitosta näiden kahden pelaajan välillä.", 0, "Onnenpyörä")
    Randomize
    Arvopa = Int((2) * Rnd + 1)
    
    Select Case Arvopa
    
    Case 1
        vteksti = "Voittaja on pelaaja 2 eli " + lblPelaaja2.Caption + "."
        BonukseenMenija = 2
    
    Case 2
        vteksti = "Voittaja on pelaaja 3 eli " + lblPelaaja3.Caption + "."
        BonukseenMenija = 3
    
    End Select
    
    msg = MsgBox(vteksti, 0, "Onnenpyörä")

Case 6
    
    msg = MsgBox("Koska kaikilla kolmella pelaajalla on sama pistemäärä suoritetaan arvonta voitosta näiden pelaajien välillä.", 0, "Onnenpyörä")
    Randomize
    Arvopa = Int((3) * Rnd + 1)
    
    Select Case Arvopa
    
    Case 1
        vteksti = "Voittaja on pelaaja 1 eli " + lblPelaaja1.Caption + "."
        BonukseenMenija = 1
    
    Case 2
        vteksti = "Voittaja on pelaaja 2 eli " + lblPelaaja2.Caption + "."
        BonukseenMenija = 2
    
    Case 3
        vteksti = "Voittaja on pelaaja 3 eli " + lblPelaaja3.Caption + "."
        BonukseenMenija = 3
    
    End Select
    
    msg = MsgBox(vteksti, 0, "Onnenpyörä")
        
End Select

Select Case BonukseenMenija

Case 1
    lblValmistaudu.Caption = frmMain.lblP1.Caption + ", valmistaudu " + SBonus

Case 2
    lblValmistaudu.Caption = frmMain.lblP2.Caption + ", valmistaudu " + SBonus

Case 3
    lblValmistaudu.Caption = frmMain.lblP3.Caption + ", valmistaudu " + SBonus

End Select

End Sub

Private Sub cmdOK_Click()

If frmRekisterointi.Rekisteroity = 1 Then
    
    frmKirjaimet.Bonus = 1
    frmKirjaimet.SB = SuperBonus
    frmMain.BonusTehtava
    Unload frmLoppupisteet

Else

    Dim msg As Integer
    msg = MsgBox("Bonustehtävään pääset vain rekisteröidyssä versiossa!", vbOKOnly, "Onnenpyörä")

    msg = MsgBox("Haluatko pelata uudelleen?", vbYesNo, "Onnenpyörä")
    
    If msg = 6 Then
        frmNapit.Uusipeli
    
    Else
        End
    
    End If

End If

End Sub

Private Sub Form_Load()

frmNapit.Enabled = False

End Sub
