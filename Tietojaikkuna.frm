VERSION 5.00
Begin VB.Form Tietojaikkuna 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tietoja: Onnenpyörä"
   ClientHeight    =   2355
   ClientLeft      =   1890
   ClientTop       =   1800
   ClientWidth     =   5070
   Icon            =   "Tietojaikkuna.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2355
   ScaleWidth      =   5070
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cdmOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   360
      Left            =   3840
      TabIndex        =   0
      Top             =   1890
      Width           =   1120
   End
   Begin VB.Label lblNimi 
      Height          =   435
      Left            =   1680
      TabIndex        =   3
      Top             =   1440
      UseMnemonic     =   0   'False
      Width           =   3285
      Visible         =   0   'False
   End
   Begin VB.Label lblOikeus 
      AutoSize        =   -1  'True
      Caption         =   "Tuotteen käyttöoikeus:"
      Height          =   195
      Left            =   1680
      TabIndex        =   2
      Top             =   1200
      Width           =   1635
      Visible         =   0   'False
   End
   Begin VB.Image Image1 
      Height          =   810
      Left            =   120
      Picture         =   "Tietojaikkuna.frx":000C
      Top             =   1440
      Width           =   1500
   End
   Begin VB.Line lnViiva2 
      BorderColor     =   &H00E0E0E0&
      X1              =   1065
      X2              =   4825
      Y1              =   1095
      Y2              =   1095
   End
   Begin VB.Line lnViiva 
      BorderColor     =   &H00808080&
      X1              =   1065
      X2              =   4825
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblCopyright 
      Caption         =   $"Tietojaikkuna.frx":3F96
      Height          =   735
      Left            =   1120
      TabIndex        =   1
      Top             =   255
      Width           =   3135
   End
   Begin VB.Image imgKuva 
      Height          =   480
      Left            =   120
      Picture         =   "Tietojaikkuna.frx":4025
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "Tietojaikkuna"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cdmOK_Click()

frmNapit.Enabled = True
Unload Tietojaikkuna

End Sub

Private Sub Form_Load()

frmNapit.Enabled = False
If frmRekisterointi.Rekisteroity = 1 Then
    lblNimi = frmRekisterointi.RekisteriNimi

Else
    lblNimi = "Rekisteröimätön kopio"

End If

End Sub


Private Sub Form_Unload(Cancel As Integer)

frmNapit.Enabled = True
Unload Tietojaikkuna

End Sub


