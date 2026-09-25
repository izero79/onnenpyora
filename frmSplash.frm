VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4260
   ClientLeft      =   4080
   ClientTop       =   3930
   ClientWidth     =   7395
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   4260
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrAika 
      Interval        =   2000
      Left            =   1680
      Top             =   3720
   End
   Begin VB.Label lblRekisteroity 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Rekisteröimätön kopio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   960
      UseMnemonic     =   0   'False
      Width           =   2415
      Visible         =   0   'False
   End
   Begin VB.Label lblVersio 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Freeeware Versio 1.3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   3750
      Width           =   3135
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private AlkuAika As Integer

Private Sub Form_Click()

Unload frmSplash
frmKirjautuminen.Show

End Sub

Private Sub Form_Load()

frmRekisterointi.LueRekTiedot

If frmRekisterointi.Rekisteroity = 1 Then
    lblRekisteroity.Caption = frmRekisterointi.RekisteriNimi
'    lblVersio.Width = 1575
'    lblVersio.Left = 5280
    lblVersio.Width = 3135
    lblVersio.Left = 3720
    lblVersio.Caption = "Freeware Versio 1.3"

Else
    lblRekisteroity.Caption = "Rekisteröimätön kopio"
    lblVersio.Width = 3135
    lblVersio.Left = 3720
    lblVersio.Caption = "Freeware Versio 1.3"

End If
Load frmKirjautuminen
AlkuAika = 0

End Sub


Private Sub lblRekisteroity_Click()

Unload frmSplash
frmKirjautuminen.Show

End Sub

Private Sub lblVersio_Click()

Unload frmSplash
frmKirjautuminen.Show

End Sub


Private Sub tmrAika_Timer()

If AlkuAika = 2 Then
    tmrAika.Enabled = False
    Unload frmSplash
    frmKirjautuminen.Show

End If
AlkuAika = AlkuAika + 1

End Sub


