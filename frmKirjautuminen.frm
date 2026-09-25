VERSION 5.00
Object = "{B9D22273-0C24-101B-AEBD-04021C009402}#1.0#0"; "KEYSTA32.OCX"
Begin VB.Form frmKirjautuminen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Onnenpyörä - Kirjautuminen"
   ClientHeight    =   3000
   ClientLeft      =   2715
   ClientTop       =   2910
   ClientWidth     =   4125
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3000
   ScaleWidth      =   4125
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRekisteroi 
      Caption         =   "&Rekisteröi..."
      Enabled         =   0   'False
      Height          =   495
      Left            =   1440
      TabIndex        =   11
      Top             =   2400
      Width           =   1215
      Visible         =   0   'False
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   2880
      TabIndex        =   5
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdLopeta 
      Cancel          =   -1  'True
      Caption         =   "&Lopeta"
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox txtNimi3 
      Height          =   375
      Left            =   1080
      MaxLength       =   25
      TabIndex        =   4
      Top             =   1800
      Width           =   2055
      Visible         =   0   'False
   End
   Begin VB.TextBox txtNimi2 
      Height          =   375
      Left            =   1080
      MaxLength       =   25
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
      Visible         =   0   'False
   End
   Begin VB.TextBox txtNimi1 
      Height          =   375
      Left            =   1080
      MaxLength       =   25
      TabIndex        =   2
      Top             =   840
      Width           =   2055
   End
   Begin VB.Frame framPelaajat 
      Caption         =   "Pelaajien lukumäärä ja nimet (max 25 merk.)"
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   3855
      Begin VB.OptionButton optPelaajat 
         Alignment       =   1  'Right Justify
         Caption         =   "&3"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   9
         Top             =   1300
         Width           =   500
      End
      Begin VB.OptionButton optPelaajat 
         Alignment       =   1  'Right Justify
         Caption         =   "&2"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   800
         Width           =   500
      End
      Begin VB.OptionButton optPelaajat 
         Alignment       =   1  'Right Justify
         Caption         =   "&1"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   300
         Value           =   -1  'True
         Width           =   500
      End
   End
   Begin KeyStatLib.MhState mhCAPS 
      Height          =   255
      Left            =   3840
      TabIndex        =   10
      Top             =   2880
      Width           =   375
      Visible         =   0   'False
      _Version        =   65536
      _ExtentX        =   661
      _ExtentY        =   450
      _StockProps     =   65
      MouseIcon       =   "frmKirjautuminen.frx":0000
   End
   Begin VB.Label lblTerve 
      AutoSize        =   -1  'True
      Caption         =   "Tervetuloa Onnenpyörän maailmaan!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3840
   End
End
Attribute VB_Name = "frmKirjautuminen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Nimi1 As String
Public Nimi2 As String
Public Nimi3 As String
Public Pluku As Integer
Private Sub Lue()

Dim A As String
Dim B As String
Dim C As String
Dim D As Integer
Dim E As Integer
Dim F As Integer
Dim tiedosto As String

tiedosto = App.Path + "\pelaajat.dat"

Open tiedosto For Random As #1

    Get #1, , A
    Get #1, , B
    Get #1, , C
    Get #1, , D
    Get #1, , E
    Get #1, , F

Close #1

txtNimi1.Text = A
txtNimi2.Text = B
txtNimi3.Text = C
optPelaajat(0).Value = D
optPelaajat(1).Value = E
optPelaajat(2).Value = F

End Sub

Private Sub Talleta()

Dim A As String
Dim B As String
Dim C As String
Dim D As Integer
Dim E As Integer
Dim F As Integer
Dim tiedosto As String

A = txtNimi1.Text
B = txtNimi2.Text
C = txtNimi3.Text
D = optPelaajat(0).Value
E = optPelaajat(1).Value
F = optPelaajat(2).Value
tiedosto = App.Path + "\pelaajat.dat"

Open tiedosto For Random As #1
    
    Put #1, , A
    Put #1, , B
    Put #1, , C
    Put #1, , D
    Put #1, , E
    Put #1, , F

Close #1

End Sub

Private Sub cmdLopeta_Click()

End

End Sub



Private Sub cmdOK_Click()
    
    Select Case Pluku

    Case 1
        Nimi1 = txtNimi1.Text
        Nimi2 = ""
        Nimi3 = ""
    Case 2
        Nimi1 = txtNimi1.Text
        Nimi2 = txtNimi2.Text
        Nimi3 = ""
    Case 3
        Nimi1 = txtNimi1.Text
        Nimi2 = txtNimi2.Text
        Nimi3 = txtNimi3.Text

    End Select

Dim msg As Integer

If mhCAPS.Value = False Then
    mhCAPS.Value = True

End If

msg = MsgBox("Muista pitää Caps Lock päällä.", 0, "Onnenpyörä")
Talleta
Load frmMain
frmNapit.Show

End Sub

Private Sub cmdRekisteroi_Click()

frmRekisterointi.Mista = 1
frmRekisterointi.Show
frmKirjautuminen.Enabled = False

End Sub

Private Sub Form_Load()

If frmRekisterointi.Rekisteroity = 1 Then
    cmdRekisteroi.Visible = False
End If

Lue

If optPelaajat(2).Value = True Then
    Pluku = 3
    NimiTarkistus
    Exit Sub

End If

If optPelaajat(1).Value = True Then
    Pluku = 2
    txtNimi3.Visible = False
    NimiTarkistus
    Exit Sub

End If

If optPelaajat(0).Value = True Then
    Pluku = 1
    txtNimi2.Visible = False
    txtNimi3.Visible = False
    NimiTarkistus
    
    Exit Sub

End If

End Sub

Private Sub optPelaajat_Click(Index As Integer)

Select Case Index

Case 0
    txtNimi2.Visible = False
    txtNimi3.Visible = False
    Pluku = 1

Case 1
    txtNimi2.Visible = True
    txtNimi3.Visible = False
    Pluku = 2

Case 2
    txtNimi2.Visible = True
    txtNimi3.Visible = True
    Pluku = 3

End Select
    
NimiTarkistus

End Sub



Private Sub NimiTarkistus()

If txtNimi1.Text <> "" And Pluku = 1 Then
    cmdOK.Enabled = True

Else
    
    If txtNimi1.Text <> "" And txtNimi2.Text <> "" And Pluku = 2 Then
        cmdOK.Enabled = True
    
    Else
        
        If txtNimi1.Text <> "" And txtNimi2.Text <> "" And txtNimi3.Text <> "" And Pluku = 3 Then
            cmdOK.Enabled = True
        
        Else
            cmdOK.Enabled = False
        
        End If
    
    End If

End If

End Sub


Private Sub optPelaajat_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

NimiTarkistus

End Sub

Private Sub txtNimi1_Change()

NimiTarkistus
        
End Sub


Private Sub txtNimi1_Click()
    
NimiTarkistus

End Sub


Private Sub txtNimi1_GotFocus()

txtNimi1.SelStart = 0
txtNimi1.SelLength = 40

End Sub

Private Sub txtNimi2_Change()
    
NimiTarkistus

End Sub


Private Sub txtNimi2_GotFocus()

txtNimi2.SelStart = 0
txtNimi2.SelLength = 40

End Sub


Private Sub txtNimi3_Change()
    
NimiTarkistus

End Sub


Private Sub txtNimi3_GotFocus()

txtNimi3.SelStart = 0
txtNimi3.SelLength = 40

End Sub


