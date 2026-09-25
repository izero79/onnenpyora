VERSION 5.00
Begin VB.Form frmRekisterointi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Onnenpyörän rekisteröinti"
   ClientHeight    =   6315
   ClientLeft      =   1665
   ClientTop       =   705
   ClientWidth     =   5640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   5640
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPeruuta 
      Caption         =   "&Peruuta"
      Height          =   495
      Left            =   4320
      TabIndex        =   12
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   4320
      TabIndex        =   11
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtKoodi 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   5760
      Width           =   3375
   End
   Begin VB.TextBox txtNimi 
      Height          =   375
      Left            =   120
      MaxLength       =   30
      TabIndex        =   7
      Top             =   5040
      Width           =   3375
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "tesisoftware@usa.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   200
      Left            =   300
      MouseIcon       =   "frmRekisterointi.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   1755
      Width           =   1580
   End
   Begin VB.Label lblLinkki 
      BackStyle       =   0  'Transparent
      Caption         =   "http://members.xoom.com/tesisoftware"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   200
      Left            =   120
      MouseIcon       =   "frmRekisterointi.frx":0152
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   4440
      Width           =   2780
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Rekisteröinti koodi:"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   5520
      Width           =   1350
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Nimi:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   4800
      Width           =   345
   End
   Begin VB.Label Label7 
      Caption         =   "Lähetetyistä tehtävistä kokoan paketteja joita voi kopioida osoitteesta"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   4080
      Width           =   3855
   End
   Begin VB.Label Label6 
      Caption         =   $"frmRekisterointi.frx":02A4
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   3855
   End
   Begin VB.Label Label5 
      Caption         =   $"frmRekisterointi.frx":0339
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   $"frmRekisterointi.frx":042E
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "1) Tee 8 normaalia ja 2 bonus tehtävää Onnenpyörän       tehtävätekijällä."
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "Tee näin:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Onnenpyörän rekisteröinti on helppoa ja ennenkaikkea ILMAINEN. Tarvitset ainoastaan käyttää vähän mielikuvitustasi ja sähköpostia."
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "frmRekisterointi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Nimi As String
Private k1 As Long
Private k2 As Long
Private k3 As Long
Private k4 As Long
Private k5 As Long
Private k6 As Long
Private k7 As Long
Private k8 As Long
Private k9 As Long
Private k10 As Long
Private k11 As Long
Private k12 As Long
Private k13 As Long
Private k14 As Long
Private k15 As Long
Private k16 As Long
Private k17 As Long
Private k18 As Long
Private k19 As Long
Private k20 As Long
Private k21 As Long
Private k22 As Long
Private k23 As Long
Private k24 As Long
Private k25 As Long
Private k26 As Long
Private k27 As Long
Private k28 As Long
Private k29 As Long
Private k30 As Long
Private pituus As Integer
Private vaille As Integer
Private code As Long
Private Yrityskoodi As Long
Public Rekisteroity As Integer
Public RekisteriNimi As String
Private UusiNimi As String
Private Talletetaanko As Integer
Public Mista As Integer

Private Sub TalletaRekTiedot()

Dim RekTiedosto As String
Dim Kirjoitettava1 As String
Dim Kirjoitettava2 As Long

RekTiedosto = App.Path + "\Reg.dat"

Kirjoitettava1 = txtNimi.Text
Kirjoitettava2 = code * 9

Open RekTiedosto For Random As #1

    Put #1, , Kirjoitettava1
    Put #1, , Kirjoitettava2

Close #1


End Sub

Public Sub LueRekTiedot()

Dim RekTiedosto As String
Dim Luettava1 As String
Dim Luettava2 As Long

RekTiedosto = App.Path + "\Reg.dat"

Open RekTiedosto For Random As #1

    Get #1, , Luettava1
    Get #1, , Luettava2

Close #1

RekisteriNimi = Luettava1
Yrityskoodi = Luettava2 / 9

Talletetaanko = 0
Tarkista

If Yrityskoodi = code Then
    Rekisteroity = 1

Else
    Rekisteroity = 1

End If

End Sub

Private Sub Tarkista()

Dim i As Integer

pituus = Len(RekisteriNimi)

If pituus < 30 Then
    vaille = 30 - pituus
    UusiNimi = RekisteriNimi
    
    For i = 1 To vaille
        UusiNimi = UusiNimi + " "
    
    Next

Else
    UusiNimi = RekisteriNimi

End If

k1 = 0
k2 = 0
k3 = 0
k4 = 0
k5 = 0
k6 = 0
k7 = 0
k8 = 0
k9 = 0
k10 = 0
k11 = 0
k12 = 0
k13 = 0
k14 = 0
k15 = 0
k16 = 0
k17 = 0
k18 = 0
k19 = 0
k20 = 0
k21 = 0
k22 = 0
k23 = 0
k24 = 0
k25 = 0
k26 = 0
k27 = 0
k28 = 0
k29 = 0
k30 = 0

k1 = Asc(Mid(UusiNimi, 1, 1))
k2 = Asc(Mid(UusiNimi, 2, 1))
k3 = Asc(Mid(UusiNimi, 3, 1))
k4 = Asc(Mid(UusiNimi, 4, 1))
k5 = Asc(Mid(UusiNimi, 5, 1))
k6 = Asc(Mid(UusiNimi, 6, 1))
k7 = Asc(Mid(UusiNimi, 7, 1))
k8 = Asc(Mid(UusiNimi, 8, 1))
k9 = Asc(Mid(UusiNimi, 9, 1))
k10 = Asc(Mid(UusiNimi, 10, 1))
k11 = Asc(Mid(UusiNimi, 11, 1))
k12 = Asc(Mid(UusiNimi, 12, 1))
k13 = Asc(Mid(UusiNimi, 13, 1))
k14 = Asc(Mid(UusiNimi, 14, 1))
k15 = Asc(Mid(UusiNimi, 15, 1))
k16 = Asc(Mid(UusiNimi, 16, 1))
k17 = Asc(Mid(UusiNimi, 17, 1))
k18 = Asc(Mid(UusiNimi, 18, 1))
k19 = Asc(Mid(UusiNimi, 19, 1))
k20 = Asc(Mid(UusiNimi, 20, 1))
k21 = Asc(Mid(UusiNimi, 21, 1))
k22 = Asc(Mid(UusiNimi, 22, 1))
k23 = Asc(Mid(UusiNimi, 23, 1))
k24 = Asc(Mid(UusiNimi, 24, 1))
k25 = Asc(Mid(UusiNimi, 25, 1))
k26 = Asc(Mid(UusiNimi, 26, 1))
k27 = Asc(Mid(UusiNimi, 27, 1))
k28 = Asc(Mid(UusiNimi, 28, 1))
k29 = Asc(Mid(UusiNimi, 29, 1))
k30 = Asc(Mid(UusiNimi, 30, 1))

code = (k1 * k2) + (k3 * k4) + (k9 * k10) + (k15 * k16) + (k20 * k21) + (k26 * k27) + k5 + k6 + k7 + k8 + k11 + k12 + k13 + k14 + k17 + k18 + k19 + k22 + k23 + k24 + k25 + k28 + k29 + k30

End Sub

Private Sub cmdOK_Click()

Dim msg As String
Talletetaanko = 1

If txtKoodi.Text = "" Then
    msg = MsgBox("Rekisteröintikoodi on väärä, tarkista ja yritä uudelleen", vbOKOnly, "Onnenpyörä")
    Exit Sub

Else
    On Error GoTo virhe
    RekisteriNimi = txtNimi.Text
    Yrityskoodi = txtKoodi.Text
    
    If Len(RekisteriNimi) < 4 Then
        msg = MsgBox("Rekisteröintikoodi on väärä, tarkista ja yritä uudelleen", vbOKOnly, "Onnenpyörä")
        Exit Sub
    
    End If
    Talletetaanko = 1
    Tarkista
    
    If Yrityskoodi = 0 Then
        msg = MsgBox("Rekisteröintikoodi on väärä, tarkista ja yritä uudelleen", vbOKOnly, "Onnenpyörä")
        Exit Sub
    
    End If
    
    If Yrityskoodi = code Then
        Rekisteroity = 1
        
        If Talletetaanko = 1 Then
            msg = MsgBox("Kiitos rekisteröinnistä. Käynnistä peli uudelleen, jotta saat pelistä kaiken irti.", vbOKOnly, "Onnenpyörä")
            TalletaRekTiedot
        
        End If
    
    Else
        Rekisteroity = 1
        msg = MsgBox("Rekisteröintikoodi on väärä, tarkista ja yritä uudelleen", vbOKOnly, "Onnenpyörä")
        Exit Sub
    
    End If

End If
    
Select Case Mista

Case 1
    
    frmKirjautuminen.Enabled = True
    frmRekisterointi.Hide
        
Case 2
    frmNapit.Enabled = True
    frmMain.Enabled = False
    frmRekisterointi.Hide
    
End Select


Exit Sub
virhe:
        msg = MsgBox("Rekisteröintikoodi on väärä, tarkista ja yritä uudelleen", vbOKOnly, "Onnenpyörä")

End Sub

Private Sub cmdPeruuta_Click()

Select Case Mista

Case 1
    
    frmKirjautuminen.Enabled = True
    frmRekisterointi.Hide
        
Case 2
    frmNapit.Enabled = True
    frmMain.Enabled = False
    frmRekisterointi.Hide
    
End Select

End Sub


Private Sub Label10_Click()

Dim i As Long
Dim pathi As String

pathi = App.Path
i = ShellExecute(Me.hWnd, vbNullString, "mailto:tesisoftware@usa.net", vbNullString, pathi, SW_SHOWNORMAL)

End Sub


Private Sub lblLinkki_Click()

Dim i As Long
Dim pathi As String

pathi = App.Path
i = ShellExecute(Me.hWnd, vbNullString, "http://members.xoom.com/tesisoftware", vbNullString, pathi, SW_SHOWNORMAL)

End Sub


