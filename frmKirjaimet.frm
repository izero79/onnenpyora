VERSION 5.00
Begin VB.Form frmKirjaimet 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Onnenpyörä - Kirjaimet"
   ClientHeight    =   3510
   ClientLeft      =   4185
   ClientTop       =   3240
   ClientWidth     =   4605
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3510
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrAika2 
      Interval        =   1000
      Left            =   2160
      Top             =   0
   End
   Begin VB.Timer tmrAika 
      Interval        =   1000
      Left            =   2400
      Top             =   720
   End
   Begin VB.CommandButton cmdKonso 
      Caption         =   "Z"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   19
      Left            =   1800
      TabIndex        =   19
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdKonso 
      Caption         =   "X"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   18
      Left            =   1320
      TabIndex        =   18
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdKonso 
      Caption         =   "W"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   17
      Left            =   840
      TabIndex        =   17
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdKonso 
      Caption         =   "V"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   360
      TabIndex        =   16
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdKonso 
      Caption         =   "T"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   1800
      TabIndex        =   15
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton cmdKonso 
      Caption         =   "S"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   14
      Left            =   1320
      TabIndex        =   14
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton cmdKonso 
      Caption         =   "R"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   13
      Left            =   840
      TabIndex        =   13
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton cmdKonso 
      Caption         =   "Q"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   12
      Left            =   360
      TabIndex        =   12
      Top             =   1800
      Width           =   375
   End
   Begin VB.CommandButton cmdKonso 
      Caption         =   "P"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   1800
      TabIndex        =   11
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdKonso 
      Caption         =   "N"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   10
      Left            =   1320
      TabIndex        =   10
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdKonso 
      Caption         =   "M"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   9
      Left            =   840
      TabIndex        =   9
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdKonso 
      Caption         =   "L"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   360
      TabIndex        =   8
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdKonso 
      Caption         =   "K"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   1800
      TabIndex        =   7
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmdKonso 
      Caption         =   "J"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   1320
      TabIndex        =   6
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmdKonso 
      Caption         =   "H"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   840
      TabIndex        =   5
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmdKonso 
      Caption         =   "G"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmdKonso 
      Caption         =   "F"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   1800
      TabIndex        =   3
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdKonso 
      Caption         =   "D"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdKonso 
      Caption         =   "C"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdKonso 
      Caption         =   "B"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdVokaali 
      Caption         =   "Ö"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   8
      Left            =   3840
      TabIndex        =   28
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdVokaali 
      Caption         =   "Ä"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   7
      Left            =   3360
      TabIndex        =   27
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdVokaali 
      Caption         =   "Å"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   6
      Left            =   2880
      TabIndex        =   26
      Top             =   1320
      Width           =   375
   End
   Begin VB.CommandButton cmdVokaali 
      Caption         =   "Y"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   3840
      TabIndex        =   25
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmdVokaali 
      Caption         =   "U"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3360
      TabIndex        =   24
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmdVokaali 
      Caption         =   "O"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2880
      TabIndex        =   23
      Top             =   840
      Width           =   375
   End
   Begin VB.CommandButton cmdVokaali 
      Caption         =   "I"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3840
      TabIndex        =   22
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdVokaali 
      Caption         =   "E"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   21
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton cmdVokaali 
      Caption         =   "A"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2880
      TabIndex        =   20
      Top             =   360
      Width           =   375
   End
   Begin VB.Frame framKonsonantit 
      Caption         =   "Konsonantit"
      Height          =   2775
      Left            =   120
      TabIndex        =   29
      Top             =   0
      Width           =   2295
   End
   Begin VB.Frame framVokaalit 
      Caption         =   "Vokaalit"
      Height          =   1815
      Left            =   2640
      TabIndex        =   30
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "Ö"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   28
      Left            =   3240
      TabIndex        =   64
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "Ä"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   27
      Left            =   3000
      TabIndex        =   63
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "Å"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   26
      Left            =   2760
      TabIndex        =   62
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "Y"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   25
      Left            =   2280
      TabIndex        =   61
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   24
      Left            =   1320
      TabIndex        =   60
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   23
      Left            =   3480
      TabIndex        =   59
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   22
      Left            =   2040
      TabIndex        =   58
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   21
      Left            =   1080
      TabIndex        =   57
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   20
      Left            =   120
      TabIndex        =   56
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "Z"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   19
      Left            =   2520
      TabIndex        =   55
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   18
      Left            =   2040
      TabIndex        =   54
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "W"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   17
      Left            =   1800
      TabIndex        =   53
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "V"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   16
      Left            =   1560
      TabIndex        =   52
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "T"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   15
      Left            =   1080
      TabIndex        =   51
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   14
      Left            =   840
      TabIndex        =   50
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   600
      TabIndex        =   49
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "Q"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   12
      Left            =   360
      TabIndex        =   48
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "P"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   47
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "N"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   3240
      TabIndex        =   46
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   3000
      TabIndex        =   45
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   2760
      TabIndex        =   44
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "K"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   2520
      TabIndex        =   43
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "J"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   2280
      TabIndex        =   42
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   41
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1560
      TabIndex        =   40
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   1320
      TabIndex        =   39
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   840
      TabIndex        =   38
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   37
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblValitut 
      Alignment       =   2  'Center
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   36
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label lblX 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3120
      TabIndex        =   35
      Top             =   2040
      Width           =   120
   End
   Begin VB.Label lblPisteet 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2760
      TabIndex        =   34
      Top             =   2400
      Width           =   675
   End
   Begin VB.Label lblKirjain 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
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
      Height          =   360
      Left            =   3315
      TabIndex        =   33
      Top             =   2000
      Width           =   420
   End
   Begin VB.Label lblMaara 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2760
      TabIndex        =   32
      Top             =   2040
      Width           =   270
   End
   Begin VB.Label lblMusta 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   2640
      TabIndex        =   31
      Top             =   1920
      Width           =   1815
   End
End
Attribute VB_Name = "frmKirjaimet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public KAika As Integer
Public Tehtavateksti As String
Public Tehtavarivi As String
Public Teksti As Variant
Public Teksti2 As Variant
Public Teksti3 As Variant
Public Teksti4 As Variant
Public k1 As String
Public k2 As String
Public k3 As String
Public k4 As String
Public k5 As String
Public k6 As String
Public k7 As String
Public k8 As String
Public k9 As String
Public k10 As String
Public k11 As String
Public k12 As String
Public k13 As String
Public A As Integer
Public B As Integer
Public C As Integer
Public D As Integer
Public E As Integer
Public F As Integer
Public G As Integer
Public H As Integer
Public II As Integer
Public J As Integer
Public K As Integer
Public L As Integer
Public M As Integer
Public N As Integer
Public O As Integer
Public P As Integer
Public Q As Integer
Public R As Integer
Public S As Integer
Public T As Integer
Public U As Integer
Public V As Integer
Public W As Integer
Public X As Integer
Public Y As Integer
Public Z As Integer
Public AO As Integer
Public AE As Integer
Public OE As Integer
Public Apu1 As String
Public Apu2 As String
Public Apu3 As String
Public SB As Integer
Public Valittujakirjaimia As Integer
Public Bonus As Integer
Public Valiaika As Integer
Public BSanottu As Integer
Public KSanottu As Integer
Private Sub Avaa_Tehtava(TehtavaNumero As Integer)

Dim Tekstirivi As String
Dim tiedosto As String

tiedosto = App.Path + "\teht.txt"

Open tiedosto For Input As #1

    Do While Not EOF(1)
    
    Line Input #1, Tekstirivi
        frmMain.numero = Left(Tekstirivi, 3)
        
        If TehtavaNumero = Int(frmMain.numero) Then
            Exit Do
        End If
    
    Loop
    
    Tehtavarivi = Tekstirivi
    Tehtavateksti = Mid(Tekstirivi, 78, 52)

Close #1

Tyhjaa_Ruudut

End Sub

Private Sub Avaa_BonusTehtava(BonusTehtavaNumero As Integer)

Dim Tekstirivi As String
Dim numero As String
Dim tiedosto As String

tiedosto = App.Path + "\bteht.txt"

Open tiedosto For Input As #1

    Do While Not EOF(1)
        Line Input #1, Tekstirivi
            numero = Left(Tekstirivi, 3)
            If BonusTehtavaNumero = Int(numero) Then
                Exit Do
            End If
    Loop
    
    Tehtavarivi = Tekstirivi
    Tehtavateksti = Mid(Tekstirivi, 84, 52)
    Apu1 = Mid(Tekstirivi, 78, 1)
    Apu2 = Mid(Tekstirivi, 80, 1)
    Apu3 = Mid(Tekstirivi, 82, 1)

Close #1

Tyhjaa_Ruudut

End Sub
Public Sub Avut()

Dim val As String

For i = 0 To 12
    val = Teksti(i)
    
    Select Case val
    
    Case Apu1
            frmTehtava.txtKirjain(i).Text = Apu1
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
    Case Apu2
            frmTehtava.txtKirjain(i).Text = Apu2
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
    Case Apu3
            frmTehtava.txtKirjain(i).Text = Apu3
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
    End Select
Next

For i = 13 To 25
    val = Teksti2(i - 13)
    
    Select Case val
    
    Case Apu1
            frmTehtava.txtKirjain(i).Text = Apu1
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
    Case Apu2
            frmTehtava.txtKirjain(i).Text = Apu2
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
    Case Apu3
            frmTehtava.txtKirjain(i).Text = Apu3
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
    End Select
Next

For i = 26 To 38
    val = Teksti3(i - 26)
    Select Case val
    
    Case Apu1
            frmTehtava.txtKirjain(i).Text = Apu1
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
    Case Apu2
            frmTehtava.txtKirjain(i).Text = Apu2
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
    Case Apu3
            frmTehtava.txtKirjain(i).Text = Apu3
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
    End Select
Next

For i = 39 To 51
    val = Teksti4(i - 39)
    Select Case val
    
    Case Apu1
            frmTehtava.txtKirjain(i).Text = Apu1
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
    Case Apu2
            frmTehtava.txtKirjain(i).Text = Apu2
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
    Case Apu3
            frmTehtava.txtKirjain(i).Text = Apu3
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
    End Select
Next

End Sub

Private Sub BKirjaimetpois()

Select Case Asc(Apu1)
    
    Case 66 To 68
        lblValitut((Asc(Apu1) - 66)).Visible = False
    
    Case 70 To 72
        lblValitut((Asc(Apu1) - 67)).Visible = False
    
    Case 74 To 78
        lblValitut((Asc(Apu1) - 68)).Visible = False
    
    Case 80 To 84
        lblValitut((Asc(Apu1) - 69)).Visible = False
    
    Case 86 To 88
        lblValitut((Asc(Apu1) - 70)).Visible = False
    
    Case 90
        lblValitut(19).Visible = False

End Select

Select Case Asc(Apu2)
    
    Case 66 To 68
        lblValitut((Asc(Apu2) - 66)).Visible = False
    
    Case 70 To 72
        lblValitut((Asc(Apu2) - 67)).Visible = False
    
    Case 74 To 78
        lblValitut((Asc(Apu2) - 68)).Visible = False
    
    Case 80 To 84
        lblValitut((Asc(Apu2) - 69)).Visible = False
    
    Case 86 To 88
        lblValitut((Asc(Apu2) - 70)).Visible = False
    
    Case 90
        lblValitut(19).Visible = False

End Select

Select Case Asc(Apu3)
    
    Case 65
        lblValitut(20).Visible = False
    
    Case 69
        lblValitut(21).Visible = False
    
    Case 73
        lblValitut(22).Visible = False
    
    Case 79
        lblValitut(23).Visible = False
    
    Case 85
        lblValitut(24).Visible = False
    
    Case 89
        lblValitut(25).Visible = False
    
    Case 197
        lblValitut(26).Visible = False
    
    Case 196
        lblValitut(27).Visible = False
    
    Case 214
        lblValitut(28).Visible = False

End Select

BKirjaimetpois2

End Sub
Private Sub BKirjaimetpois2()

Select Case Asc(Apu1)
    
    Case 66
        B = 1
    
    Case 67
        C = 1
    
    Case 68
        D = 1
    
    Case 70
        F = 1
    
    Case 71
        G = 1
    
    Case 72
        H = 1
    
    Case 74
        J = 1
    
    Case 75
        K = 1
    
    Case 76
        L = 1
    
    Case 77
        M = 1
    
    Case 78
        N = 1
    
    Case 80
        P = 1
    
    Case 81
        Q = 1
    
    Case 82
        R = 1
    
    Case 83
        S = 1
    
    Case 84
        T = 1
    
    Case 86
        V = 1
    
    Case 87
        W = 1
    
    Case 88
        X = 1
    
    Case 90
        Z = 1

End Select

Select Case Asc(Apu2)
    
    Case 66
        B = 1
    
    Case 67
        C = 1
    
    Case 68
        D = 1
    
    Case 70
        F = 1
    
    Case 71
        G = 1
    
    Case 72
        H = 1
    
    Case 74
        J = 1
    
    Case 75
        K = 1
    
    Case 76
        L = 1
    
    Case 77
        M = 1
    
    Case 78
        N = 1
    
    Case 80
        P = 1
    
    Case 81
        Q = 1
    
    Case 82
        R = 1
    
    Case 83
        S = 1
    
    Case 84
        T = 1
    
    Case 86
        V = 1
    
    Case 87
        W = 1
    
    Case 88
        X = 1
    
    Case 90
        Z = 1

End Select

Select Case Asc(Apu3)
    
    Case 65
        A = 1
    
    Case 69
        E = 1
    
    Case 73
        II = 1
    
    Case 79
        O = 1
    
    Case 85
        U = 1
    
    Case 89
        Y = 1
    
    Case 196
        AE = 1
    
    Case 197
        AO = 1
    
    Case 214
        OE = 1

End Select

End Sub

Public Sub BKonso(Index As Integer)

Dim kirjain As String

BSanottu = 0
lblValitut(Index).Visible = False
Select Case Index
    
Case 0
    
    If B = 0 Then
    B = 1
    kirjain = "B"
    
        For i = 0 To 12
        
            If Teksti(i) = "B" Then
                frmTehtava.txtKirjain(i).Text = "B"
                frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            End If
        Next
    
        For i = 13 To 25
            
            If Teksti2(i - 13) = "B" Then
                frmTehtava.txtKirjain(i).Text = "B"
                frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            End If
        Next
    
        For i = 26 To 38
        
            If Teksti3(i - 26) = "B" Then
                frmTehtava.txtKirjain(i).Text = "B"
                frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            End If
    
        Next
    
        For i = 39 To 51
        
            If Teksti4(i - 39) = "B" Then
                frmTehtava.txtKirjain(i).Text = "B"
                frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            End If
    
        Next
    
    Else
        
        BKirjainsanottu
    
    End If

Case 1
    
    If C = 0 Then
        C = 1
        kirjain = "C"
    
        For i = 0 To 12
        
            If Teksti(i) = "C" Then
                frmTehtava.txtKirjain(i).Text = "C"
                frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            End If
    
        Next
    
        For i = 13 To 25
            If Teksti2(i - 13) = "C" Then
                frmTehtava.txtKirjain(i).Text = "C"
                frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            End If
        Next
        
        For i = 26 To 38
            If Teksti3(i - 26) = "C" Then
                frmTehtava.txtKirjain(i).Text = "C"
                frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            End If
        Next
        
        For i = 39 To 51
            If Teksti4(i - 39) = "C" Then
                frmTehtava.txtKirjain(i).Text = "C"
                frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            End If
        Next
    
    Else
        
        BKirjainsanottu
    
    End If

Case 2
    If D = 0 Then
    D = 1
    kirjain = "D"
    For i = 0 To 12
        If Teksti(i) = "D" Then
            frmTehtava.txtKirjain(i).Text = "D"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "D" Then
            frmTehtava.txtKirjain(i).Text = "D"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "D" Then
            frmTehtava.txtKirjain(i).Text = "D"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "D" Then
            frmTehtava.txtKirjain(i).Text = "D"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        End If
    Next
    Else
        BKirjainsanottu
    End If

Case 3
    If F = 0 Then
    F = 1
    kirjain = "F"
    For i = 0 To 12
        If Teksti(i) = "F" Then
            frmTehtava.txtKirjain(i).Text = "F"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "F" Then
            frmTehtava.txtKirjain(i).Text = "F"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "F" Then
            frmTehtava.txtKirjain(i).Text = "F"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "F" Then
            frmTehtava.txtKirjain(i).Text = "F"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        End If
    Next
    Else
        BKirjainsanottu
    End If

Case 4
    If G = 0 Then
    G = 1
    kirjain = "G"
    For i = 0 To 12
        If Teksti(i) = "G" Then
            frmTehtava.txtKirjain(i).Text = "G"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "G" Then
            frmTehtava.txtKirjain(i).Text = "G"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "G" Then
            frmTehtava.txtKirjain(i).Text = "G"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "G" Then
            frmTehtava.txtKirjain(i).Text = "G"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    Else
        BKirjainsanottu
    End If

Case 5
    If H = 0 Then
    H = 1
    kirjain = "H"
    For i = 0 To 12
        If Teksti(i) = "H" Then
            frmTehtava.txtKirjain(i).Text = "H"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "H" Then
            frmTehtava.txtKirjain(i).Text = "H"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "H" Then
            frmTehtava.txtKirjain(i).Text = "H"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "H" Then
            frmTehtava.txtKirjain(i).Text = "H"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    Else
        BKirjainsanottu
    End If

Case 6
    If J = 0 Then
    J = 1
    kirjain = "J"
    For i = 0 To 12
        If Teksti(i) = "J" Then
            frmTehtava.txtKirjain(i).Text = "J"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "J" Then
            frmTehtava.txtKirjain(i).Text = "J"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "J" Then
            frmTehtava.txtKirjain(i).Text = "J"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "J" Then
            frmTehtava.txtKirjain(i).Text = "J"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    Else
        BKirjainsanottu
    End If

Case 7
    If K = 0 Then
    K = 1
    kirjain = "K"
    For i = 0 To 12
        If Teksti(i) = "K" Then
            frmTehtava.txtKirjain(i).Text = "K"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "K" Then
            frmTehtava.txtKirjain(i).Text = "K"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "K" Then
            frmTehtava.txtKirjain(i).Text = "K"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "K" Then
            frmTehtava.txtKirjain(i).Text = "K"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    Else
        BKirjainsanottu
    End If

Case 8
    If L = 0 Then
    L = 1
    kirjain = "L"
    For i = 0 To 12
        If Teksti(i) = "L" Then
            frmTehtava.txtKirjain(i).Text = "L"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "L" Then
            frmTehtava.txtKirjain(i).Text = "L"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "L" Then
            frmTehtava.txtKirjain(i).Text = "L"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "L" Then
            frmTehtava.txtKirjain(i).Text = "L"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    Else
        BKirjainsanottu
    End If
Case 9
    If M = 0 Then
    M = 1
    kirjain = "M"
    For i = 0 To 12
        If Teksti(i) = "M" Then
            frmTehtava.txtKirjain(i).Text = "M"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "M" Then
            frmTehtava.txtKirjain(i).Text = "M"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "M" Then
            frmTehtava.txtKirjain(i).Text = "M"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "M" Then
            frmTehtava.txtKirjain(i).Text = "M"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    Else
        BKirjainsanottu
    End If
Case 10
    If N = 0 Then
    N = 1
    kirjain = "N"
    For i = 0 To 12
        If Teksti(i) = "N" Then
            frmTehtava.txtKirjain(i).Text = "N"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "N" Then
            frmTehtava.txtKirjain(i).Text = "N"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "N" Then
            frmTehtava.txtKirjain(i).Text = "N"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "N" Then
            frmTehtava.txtKirjain(i).Text = "N"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    Else
        BKirjainsanottu
    End If
Case 11
    If P = 0 Then
    P = 1
    kirjain = "P"
    For i = 0 To 12
        If Teksti(i) = "P" Then
            frmTehtava.txtKirjain(i).Text = "P"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "P" Then
            frmTehtava.txtKirjain(i).Text = "P"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "P" Then
            frmTehtava.txtKirjain(i).Text = "P"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "P" Then
            frmTehtava.txtKirjain(i).Text = "P"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    Else
        BKirjainsanottu
    End If
Case 12
    If Q = 0 Then
    Q = 1
    kirjain = "Q"
    For i = 0 To 12
        If Teksti(i) = "Q" Then
            frmTehtava.txtKirjain(i).Text = "Q"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "Q" Then
            frmTehtava.txtKirjain(i).Text = "Q"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "Q" Then
            frmTehtava.txtKirjain(i).Text = "Q"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "Q" Then
            frmTehtava.txtKirjain(i).Text = "Q"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    Else
        BKirjainsanottu
    End If
Case 13
    If R = 0 Then
    R = 1
    kirjain = "R"
    For i = 0 To 12
        If Teksti(i) = "R" Then
            frmTehtava.txtKirjain(i).Text = "R"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "R" Then
            frmTehtava.txtKirjain(i).Text = "R"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "R" Then
            frmTehtava.txtKirjain(i).Text = "R"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
             
       End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "R" Then
            frmTehtava.txtKirjain(i).Text = "R"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
             
       End If
    Next
    Else
        BKirjainsanottu
    End If
Case 14
    If S = 0 Then
    S = 1
    kirjain = "S"
    For i = 0 To 12
        If Teksti(i) = "S" Then
            frmTehtava.txtKirjain(i).Text = "S"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "S" Then
            frmTehtava.txtKirjain(i).Text = "S"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
       End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "S" Then
            frmTehtava.txtKirjain(i).Text = "S"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "S" Then
            frmTehtava.txtKirjain(i).Text = "S"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    Else
        BKirjainsanottu
    End If
Case 15
    If T = 0 Then
    T = 1
    kirjain = "T"
    For i = 0 To 12
        If Teksti(i) = "T" Then
            frmTehtava.txtKirjain(i).Text = "T"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "T" Then
            frmTehtava.txtKirjain(i).Text = "T"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "T" Then
            frmTehtava.txtKirjain(i).Text = "T"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "T" Then
            frmTehtava.txtKirjain(i).Text = "T"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    Else
        BKirjainsanottu
    End If
Case 16
    If V = 0 Then
    V = 1
    kirjain = "V"
    For i = 0 To 12
        If Teksti(i) = "V" Then
            frmTehtava.txtKirjain(i).Text = "V"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "V" Then
            frmTehtava.txtKirjain(i).Text = "V"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "V" Then
            frmTehtava.txtKirjain(i).Text = "V"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "V" Then
            frmTehtava.txtKirjain(i).Text = "V"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    Else
        BKirjainsanottu
    End If
Case 17
    If W = 0 Then
    W = 1
    kirjain = "W"
    For i = 0 To 12
        If Teksti(i) = "W" Then
            frmTehtava.txtKirjain(i).Text = "W"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "W" Then
            frmTehtava.txtKirjain(i).Text = "W"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "W" Then
            frmTehtava.txtKirjain(i).Text = "W"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "W" Then
            frmTehtava.txtKirjain(i).Text = "W"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    Else
        BKirjainsanottu
    End If
Case 18
    If X = 0 Then
    X = 1
    kirjain = "X"
    For i = 0 To 12
        If Teksti(i) = "X" Then
            frmTehtava.txtKirjain(i).Text = "X"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "X" Then
            frmTehtava.txtKirjain(i).Text = "X"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "X" Then
            frmTehtava.txtKirjain(i).Text = "X"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "X" Then
            frmTehtava.txtKirjain(i).Text = "X"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    Else
        BKirjainsanottu
    End If
Case 19
    If Z = 0 Then
    Z = 1
    kirjain = "Z"
    For i = 0 To 12
        If Teksti(i) = "Z" Then
            frmTehtava.txtKirjain(i).Text = "Z"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "Z" Then
            frmTehtava.txtKirjain(i).Text = "Z"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "Z" Then
            frmTehtava.txtKirjain(i).Text = "Z"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "Z" Then
            frmTehtava.txtKirjain(i).Text = "Z"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    Else
        BKirjainsanottu
    End If

End Select

If BSanottu = 0 Then
    Valittujakirjaimia = Valittujakirjaimia - 1
    
    Select Case Valittujakirjaimia
    
    Case 1
        frmMain.KonsonantitPiiloon
    
    Case 0
        frmMain.KonsonantitPiiloon
        BonusAikaAlkaa
    
    End Select

End If

End Sub

Public Sub BonusAikaAlkaa()

Load frmBonusRatkaisu
frmBonusRatkaisu.SetFocus

End Sub

Public Sub BonusAloitus2()

tmrAika2.Enabled = False

If SB = 1 Then
    aputeksti = "Apuina SuperBonustehtävään annetaan konsonantit " + Apu1 + " ja " + Apu2 + ". Sekä vokaali " + Apu3
Else
    aputeksti = "Apuina Bonustehtävään annetaan konsonantit " + Apu1 + " ja " + Apu2 + ". Sekä vokaali " + Apu3
End If

msg = MsgBox(aputeksti, 0, "Onnenpyörä")
Avut
BKirjaimetpois
ValitaanKirjaimet
frmKirjaimet.SetFocus

End Sub

Public Sub BonusRatkaistu()

frmBonusRatkaisu.Hide
msg = MsgBox("Onneksi olkoon se on oikein!", 0, "Onnenpyörä")
frmBPalkinnot.optVariValinta(0).Visible = False
frmBPalkinnot.optVariValinta(1).Visible = False
frmBPalkinnot.optVariValinta(2).Visible = False
frmBPalkinnot.lblValitse.Visible = False
frmBPalkinnot.lblPalkintoTeksti(0).Top = 1000
frmBPalkinnot.lblPalkintoTeksti(0).Left = 1200
frmBPalkinnot.lblPalkinnot(0).Visible = False
frmBPalkinnot.lblPalkinnot(1).Visible = False
frmBPalkinnot.lblPalkinnot(2).Visible = False
frmBPalkinnot.pboxKuva(0).Visible = False
frmBPalkinnot.pboxKuva(1).Visible = False
frmBPalkinnot.cmdOK.Visible = False
frmBPalkinnot.cmdOK2.Visible = True

Randomize
Palkinto = Int((3) * Rnd + 1)
'
Select Case Palkinto

Case 1
    frmBPalkinnot.pboxKuva(2).Picture = frmPalkinto.pcPalkinnot.GraphicCell(frmMain.BPalkinnonnumero1)
    frmBPalkinnot.lblPalkintoTeksti(1).Visible = False
    frmBPalkinnot.lblPalkintoTeksti(2).Visible = False
    frmBPalkinnot.lblPalkintoTeksti(0).Top = 1000
    frmBPalkinnot.lblPalkintoTeksti(0).Left = 1200
    frmBPalkinnot.lblPalkintoTeksti(0).Height = 2000
    Select Case frmLoppupisteet.BonukseenMenija
    
    Case 1
        frmMain.RahaMaaraP1 = frmMain.RahaMaaraP1 + frmBPalkinnot.Arvo1
    Case 2
        frmMain.RahaMaaraP2 = frmMain.RahaMaaraP2 + frmBPalkinnot.Arvo1
    Case 3
        frmMain.RahaMaaraP3 = frmMain.RahaMaaraP3 + frmBPalkinnot.Arvo1
    
    End Select
        
    Select Case frmMain.PalkintoVari
    
    Case 1
        frmBPalkinnot.lblPalkintoTeksti(0).Caption = "Nyt katsotaan sitten mitä voitit. Punaisesta kuoresta löytyy... " + frmMain.BPalkinnonteksti1
    
    Case 2
        frmBPalkinnot.lblPalkintoTeksti(0).Caption = "Nyt katsotaan sitten mitä voitit. Keltaisesta kuoresta löytyy... " + frmMain.BPalkinnonteksti1
    
    Case 3
        frmBPalkinnot.lblPalkintoTeksti(0).Caption = "Nyt katsotaan sitten mitä voitit. Vihreästä kuoresta löytyy... " + frmMain.BPalkinnonteksti1
    
    End Select

Case 2
    frmBPalkinnot.pboxKuva(2).Picture = frmPalkinto.pcPalkinnot.GraphicCell(frmMain.BPalkinnonnumero2)
    frmBPalkinnot.lblPalkintoTeksti(0).Visible = False
    frmBPalkinnot.lblPalkintoTeksti(2).Visible = False
    frmBPalkinnot.lblPalkintoTeksti(1).Top = 1000
    frmBPalkinnot.lblPalkintoTeksti(1).Left = 1200
    frmBPalkinnot.lblPalkintoTeksti(1).Height = 2000
    
    Select Case frmLoppupisteet.BonukseenMenija
    
    Case 1
        frmMain.RahaMaaraP1 = frmMain.RahaMaaraP1 + frmBPalkinnot.Arvo2
    Case 2
        frmMain.RahaMaaraP2 = frmMain.RahaMaaraP2 + frmBPalkinnot.Arvo2
    Case 3
        frmMain.RahaMaaraP3 = frmMain.RahaMaaraP3 + frmBPalkinnot.Arvo2
    
    End Select
    
    Select Case frmMain.PalkintoVari
    
    Case 1
        frmBPalkinnot.lblPalkintoTeksti(1).Caption = "Nyt katsotaan sitten mitä voitit. Punaisesta kuoresta löytyy... " + frmMain.BPalkinnonteksti2
    
    Case 2
        frmBPalkinnot.lblPalkintoTeksti(1).Caption = "Nyt katsotaan sitten mitä voitit. Keltaisesta kuoresta löytyy... " + frmMain.BPalkinnonteksti2
    
    Case 3
        frmBPalkinnot.lblPalkintoTeksti(1).Caption = "Nyt katsotaan sitten mitä voitit. Vihreästä kuoresta löytyy... " + frmMain.BPalkinnonteksti2
    
    End Select

Case 3
    frmBPalkinnot.pboxKuva(2).Picture = frmPalkinto.pcPalkinnot.GraphicCell(frmMain.BPalkinnonnumero3)
    frmBPalkinnot.lblPalkintoTeksti(0).Visible = False
    frmBPalkinnot.lblPalkintoTeksti(1).Visible = False
    frmBPalkinnot.lblPalkintoTeksti(2).Top = 1000
    frmBPalkinnot.lblPalkintoTeksti(2).Left = 1200
    frmBPalkinnot.lblPalkintoTeksti(2).Height = 2000
    
    Select Case frmLoppupisteet.BonukseenMenija
    
    Case 1
        frmMain.RahaMaaraP1 = frmMain.RahaMaaraP1 + frmBPalkinnot.Arvo3
    Case 2
        frmMain.RahaMaaraP2 = frmMain.RahaMaaraP2 + frmBPalkinnot.Arvo3
    Case 3
        frmMain.RahaMaaraP3 = frmMain.RahaMaaraP3 + frmBPalkinnot.Arvo3
    
    End Select
    
    Select Case frmMain.PalkintoVari
    
    Case 1
        frmBPalkinnot.lblPalkintoTeksti(2).Caption = "Nyt katsotaan sitten mitä voitit. Punaisesta kuoresta löytyy... " + frmMain.BPalkinnonteksti3
    
    Case 2
        frmBPalkinnot.lblPalkintoTeksti(2).Caption = "Nyt katsotaan sitten mitä voitit. Keltaisesta kuoresta löytyy... " + frmMain.BPalkinnonteksti3
    
    Case 3
        frmBPalkinnot.lblPalkintoTeksti(2).Caption = "Nyt katsotaan sitten mitä voitit. Vihreästä kuoresta löytyy... " + frmMain.BPalkinnonteksti3
    
    End Select

End Select

frmKirjaimet.Enabled = False

If frmLoppupisteet.SuperBonus <> 1 Then
    frmBPalkinnot.Caption = "Bonuspalkinnot - Onnenpyörä"

Else
    frmBPalkinnot.Caption = "SuperBonuspalkinnot - Onnenpyörä"

End If

frmBPalkinnot.cmdOK2.Default = True
frmBPalkinnot.Show

End Sub

Public Sub BonusEiRatkaistu()

frmBonusRatkaisu.Hide
frmBPalkinnot.optVariValinta(0).Visible = False
frmBPalkinnot.optVariValinta(1).Visible = False
frmBPalkinnot.optVariValinta(2).Visible = False
frmBPalkinnot.lblValitse.Visible = False
frmBPalkinnot.lblPalkintoTeksti(0).Top = 1000
frmBPalkinnot.lblPalkintoTeksti(0).Left = 1200
frmBPalkinnot.lblPalkinnot(0).Visible = False
frmBPalkinnot.lblPalkinnot(1).Visible = False
frmBPalkinnot.lblPalkinnot(2).Visible = False
frmBPalkinnot.pboxKuva(0).Visible = False
frmBPalkinnot.pboxKuva(1).Visible = False
frmBPalkinnot.cmdOK.Visible = False
frmBPalkinnot.cmdOK2.Visible = True
Randomize
Palkinto = Int((3) * Rnd + 1)
Palkinto = 1

Select Case Palkinto

Case 1
    frmBPalkinnot.pboxKuva(2).Picture = frmPalkinto.pcPalkinnot.GraphicCell(frmMain.BPalkinnonnumero1)
    frmBPalkinnot.lblPalkintoTeksti(1).Visible = False
    frmBPalkinnot.lblPalkintoTeksti(2).Visible = False
    frmBPalkinnot.lblPalkintoTeksti(0).Top = 1000
    frmBPalkinnot.lblPalkintoTeksti(0).Left = 1200
    frmBPalkinnot.lblPalkintoTeksti(0).Height = 2000
    Select Case frmMain.PalkintoVari
    
    Case 1
        frmBPalkinnot.lblPalkintoTeksti(0).Caption = "Nyt katsotaan mitä olisit voinut voittaa. Punaisesta kuoresta löytyy... " + frmMain.BPalkinnonteksti1

    Case 2
        frmBPalkinnot.lblPalkintoTeksti(0).Caption = "Nyt katsotaan mitä olisit voinut voittaa. Keltaisesta kuoresta löytyy... " + frmMain.BPalkinnonteksti1
    
    Case 3
        frmBPalkinnot.lblPalkintoTeksti(0).Caption = "Nyt katsotaan mitä olisit voinut voittaa. Vihreästä kuoresta löytyy... " + frmMain.BPalkinnonteksti1
    
    End Select

Case 2
    frmBPalkinnot.pboxKuva(2).Picture = frmPalkinto.pcPalkinnot.GraphicCell(frmMain.BPalkinnonnumero2)
    frmBPalkinnot.lblPalkintoTeksti(0).Visible = False
    frmBPalkinnot.lblPalkintoTeksti(2).Visible = False
    frmBPalkinnot.lblPalkintoTeksti(1).Top = 1000
    frmBPalkinnot.lblPalkintoTeksti(1).Left = 1200
    frmBPalkinnot.lblPalkintoTeksti(1).Height = 2000
    
    Select Case frmMain.PalkintoVari
    
    Case 1
        frmBPalkinnot.lblPalkintoTeksti(1).Caption = "Nyt katsotaan mitä olisit voinut voittaa. Punaisesta kuoresta löytyy... " + frmMain.BPalkinnonteksti2
    
    Case 2
        frmBPalkinnot.lblPalkintoTeksti(1).Caption = "Nyt katsotaan mitä olisit voinut voittaa. Keltaisesta kuoresta löytyy... " + frmMain.BPalkinnonteksti2
    
    Case 3
        frmBPalkinnot.lblPalkintoTeksti(1).Caption = "Nyt katsotaan mitä olisit voinut voittaa. Vihreästä kuoresta löytyy... " + frmMain.BPalkinnonteksti2
    
    End Select

Case 3
    frmBPalkinnot.pboxKuva(2).Picture = frmPalkinto.pcPalkinnot.GraphicCell(frmMain.BPalkinnonnumero3)
    frmBPalkinnot.lblPalkintoTeksti(0).Visible = False
    frmBPalkinnot.lblPalkintoTeksti(1).Visible = False
    frmBPalkinnot.lblPalkintoTeksti(2).Top = 1000
    frmBPalkinnot.lblPalkintoTeksti(2).Left = 1200
    frmBPalkinnot.lblPalkintoTeksti(2).Height = 2000
    
    Select Case frmMain.PalkintoVari
    
    Case 1
        frmBPalkinnot.lblPalkintoTeksti(2).Caption = "Nyt katsotaan mitä olisit voinut voittaa. Punaisesta kuoresta löytyy... " + frmMain.BPalkinnonteksti3
    
    Case 2
        frmBPalkinnot.lblPalkintoTeksti(2).Caption = "Nyt katsotaan mitä olisit voinut voittaa. Keltaisesta kuoresta löytyy... " + frmMain.BPalkinnonteksti3
    
    Case 3
        frmBPalkinnot.lblPalkintoTeksti(2).Caption = "Nyt katsotaan mitä olisit voinut voittaa. Vihreästä kuoresta löytyy... " + frmMain.BPalkinnonteksti3
    
    End Select

End Select

frmKirjaimet.Enabled = False

If frmLoppupisteet.SuperBonus <> 1 Then
    frmBPalkinnot.Caption = "Bonuspalkinnot - Onnenpyörä"

Else
    frmBPalkinnot.Caption = "SuperBonuspalkinnot - Onnenpyörä"

End If

frmBPalkinnot.cmdOK2.Default = True
frmBPalkinnot.Show

End Sub

Private Sub Kirjainsanottu()

Dim msg

frmTehtava.Visible = False
frmKirjaimet.Visible = False
frmMain.tmrAika.Enabled = False
frmKirjaimet.tmrAika.Enabled = False
frmKirjaimet.tmrAika2.Enabled = False
msg = MsgBox("Valitsemasi kirjain oli jo valittu.", 0, "Onnenpyörä")
frmTehtava.Visible = True
frmKirjaimet.Visible = True
KSanottu = 1

End Sub

Private Sub BKirjainsanottu()

Dim msg

msg = MsgBox("Valitsemasi kirjain on jo valittu.", 0, "Onnenpyörä")
BSanottu = 1

End Sub

Public Sub Konso(Index As Integer)

Dim maara As Integer
Dim KokoPistemaara As Integer
Dim kirjain As String

KokoPistemaara = 0
KSanottu = 0
maara = 0
lblValitut(Index).Visible = False

Select Case Index
    
Case 0
    If B = 0 Then
        B = 1
        kirjain = "B"
    For i = 0 To 12
        If Teksti(i) = "B" Then
            frmTehtava.txtKirjain(i).Text = "B"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "B" Then
            frmTehtava.txtKirjain(i).Text = "B"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "B" Then
            frmTehtava.txtKirjain(i).Text = "B"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "B" Then
            frmTehtava.txtKirjain(i).Text = "B"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If
Case 1
    If C = 0 Then
    C = 1
    kirjain = "C"
    For i = 0 To 12
        If Teksti(i) = "C" Then
            frmTehtava.txtKirjain(i).Text = "C"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "C" Then
            frmTehtava.txtKirjain(i).Text = "C"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "C" Then
            frmTehtava.txtKirjain(i).Text = "C"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "C" Then
            frmTehtava.txtKirjain(i).Text = "C"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If

Case 2
    If D = 0 Then
    D = 1
    kirjain = "D"
    For i = 0 To 12
        If Teksti(i) = "D" Then
            frmTehtava.txtKirjain(i).Text = "D"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "D" Then
            frmTehtava.txtKirjain(i).Text = "D"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "D" Then
            frmTehtava.txtKirjain(i).Text = "D"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "D" Then
            frmTehtava.txtKirjain(i).Text = "D"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If

Case 3
    If F = 0 Then
    F = 1
    kirjain = "F"
    For i = 0 To 12
        If Teksti(i) = "F" Then
            frmTehtava.txtKirjain(i).Text = "F"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "F" Then
            frmTehtava.txtKirjain(i).Text = "F"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "F" Then
            frmTehtava.txtKirjain(i).Text = "F"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "F" Then
            frmTehtava.txtKirjain(i).Text = "F"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If

Case 4
    If G = 0 Then
    G = 1
    kirjain = "G"
    For i = 0 To 12
        If Teksti(i) = "G" Then
            frmTehtava.txtKirjain(i).Text = "G"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "G" Then
            frmTehtava.txtKirjain(i).Text = "G"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "G" Then
            frmTehtava.txtKirjain(i).Text = "G"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "G" Then
            frmTehtava.txtKirjain(i).Text = "G"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If

Case 5
    If H = 0 Then
    H = 1
    kirjain = "H"
    For i = 0 To 12
        If Teksti(i) = "H" Then
            frmTehtava.txtKirjain(i).Text = "H"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "H" Then
            frmTehtava.txtKirjain(i).Text = "H"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "H" Then
            frmTehtava.txtKirjain(i).Text = "H"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "H" Then
            frmTehtava.txtKirjain(i).Text = "H"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If

Case 6
    If J = 0 Then
    J = 1
    kirjain = "J"
    For i = 0 To 12
        If Teksti(i) = "J" Then
            frmTehtava.txtKirjain(i).Text = "J"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "J" Then
            frmTehtava.txtKirjain(i).Text = "J"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "J" Then
            frmTehtava.txtKirjain(i).Text = "J"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "J" Then
            frmTehtava.txtKirjain(i).Text = "J"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If

Case 7
    If K = 0 Then
    K = 1
    kirjain = "K"
    For i = 0 To 12
        If Teksti(i) = "K" Then
            frmTehtava.txtKirjain(i).Text = "K"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "K" Then
            frmTehtava.txtKirjain(i).Text = "K"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "K" Then
            frmTehtava.txtKirjain(i).Text = "K"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "K" Then
            frmTehtava.txtKirjain(i).Text = "K"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If

Case 8
    If L = 0 Then
    L = 1
    kirjain = "L"
    For i = 0 To 12
        If Teksti(i) = "L" Then
            frmTehtava.txtKirjain(i).Text = "L"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "L" Then
            frmTehtava.txtKirjain(i).Text = "L"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "L" Then
            frmTehtava.txtKirjain(i).Text = "L"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "L" Then
            frmTehtava.txtKirjain(i).Text = "L"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If
Case 9
    If M = 0 Then
    M = 1
    kirjain = "M"
    For i = 0 To 12
        If Teksti(i) = "M" Then
            frmTehtava.txtKirjain(i).Text = "M"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "M" Then
            frmTehtava.txtKirjain(i).Text = "M"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "M" Then
            frmTehtava.txtKirjain(i).Text = "M"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "M" Then
            frmTehtava.txtKirjain(i).Text = "M"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If
Case 10
    If N = 0 Then
    N = 1
    kirjain = "N"
    For i = 0 To 12
        If Teksti(i) = "N" Then
            frmTehtava.txtKirjain(i).Text = "N"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "N" Then
            frmTehtava.txtKirjain(i).Text = "N"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "N" Then
            frmTehtava.txtKirjain(i).Text = "N"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "N" Then
            frmTehtava.txtKirjain(i).Text = "N"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If
Case 11
    If P = 0 Then
    P = 1
    kirjain = "P"
    For i = 0 To 12
        If Teksti(i) = "P" Then
            frmTehtava.txtKirjain(i).Text = "P"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "P" Then
            frmTehtava.txtKirjain(i).Text = "P"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "P" Then
            frmTehtava.txtKirjain(i).Text = "P"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "P" Then
            frmTehtava.txtKirjain(i).Text = "P"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If
Case 12
    If Q = 0 Then
    Q = 1
    kirjain = "Q"
    For i = 0 To 12
        If Teksti(i) = "Q" Then
            frmTehtava.txtKirjain(i).Text = "Q"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "Q" Then
            frmTehtava.txtKirjain(i).Text = "Q"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "Q" Then
            frmTehtava.txtKirjain(i).Text = "Q"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "Q" Then
            frmTehtava.txtKirjain(i).Text = "Q"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If
Case 13
    If R = 0 Then
    R = 1
    kirjain = "R"
    For i = 0 To 12
        If Teksti(i) = "R" Then
            frmTehtava.txtKirjain(i).Text = "R"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "R" Then
            frmTehtava.txtKirjain(i).Text = "R"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "R" Then
            frmTehtava.txtKirjain(i).Text = "R"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
             maara = maara + 1
       End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "R" Then
            frmTehtava.txtKirjain(i).Text = "R"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
             maara = maara + 1
       End If
    Next
    Else
        Kirjainsanottu
    End If
Case 14
    If S = 0 Then
    S = 1
    kirjain = "S"
    For i = 0 To 12
        If Teksti(i) = "S" Then
            frmTehtava.txtKirjain(i).Text = "S"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "S" Then
            frmTehtava.txtKirjain(i).Text = "S"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
       End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "S" Then
            frmTehtava.txtKirjain(i).Text = "S"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "S" Then
            frmTehtava.txtKirjain(i).Text = "S"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If
Case 15
    If T = 0 Then
    T = 1
    kirjain = "T"
    For i = 0 To 12
        If Teksti(i) = "T" Then
            frmTehtava.txtKirjain(i).Text = "T"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "T" Then
            frmTehtava.txtKirjain(i).Text = "T"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "T" Then
            frmTehtava.txtKirjain(i).Text = "T"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "T" Then
            frmTehtava.txtKirjain(i).Text = "T"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If
Case 16
    If V = 0 Then
    V = 1
    kirjain = "V"
    For i = 0 To 12
        If Teksti(i) = "V" Then
            frmTehtava.txtKirjain(i).Text = "V"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "V" Then
            frmTehtava.txtKirjain(i).Text = "V"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "V" Then
            frmTehtava.txtKirjain(i).Text = "V"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "V" Then
            frmTehtava.txtKirjain(i).Text = "V"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If
Case 17
    If W = 0 Then
    W = 1
    kirjain = "W"
    For i = 0 To 12
        If Teksti(i) = "W" Then
            frmTehtava.txtKirjain(i).Text = "W"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "W" Then
            frmTehtava.txtKirjain(i).Text = "W"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "W" Then
            frmTehtava.txtKirjain(i).Text = "W"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "W" Then
            frmTehtava.txtKirjain(i).Text = "W"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If
Case 18
    If X = 0 Then
    X = 1
    kirjain = "X"
    For i = 0 To 12
        If Teksti(i) = "X" Then
            frmTehtava.txtKirjain(i).Text = "X"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "X" Then
            frmTehtava.txtKirjain(i).Text = "X"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "X" Then
            frmTehtava.txtKirjain(i).Text = "X"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "X" Then
            frmTehtava.txtKirjain(i).Text = "X"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If
Case 19
    If Z = 0 Then
    Z = 1
    kirjain = "Z"
    For i = 0 To 12
        If Teksti(i) = "Z" Then
            frmTehtava.txtKirjain(i).Text = "Z"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "Z" Then
            frmTehtava.txtKirjain(i).Text = "Z"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "Z" Then
            frmTehtava.txtKirjain(i).Text = "Z"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "Z" Then
            frmTehtava.txtKirjain(i).Text = "Z"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If

End Select

KokoPistemaara = maara * frmMain.SektoriPisteet
lblMaara.Caption = maara
lblKirjain.Caption = kirjain
frmMain.KonsonantitPiiloon
frmMain.VokaalitPiiloon

If KSanottu <> 1 Then
    
    If maara = 0 Then
        frmTehtava.Visible = False
        frmKirjaimet.Visible = False
        frmMain.tmrAika.Enabled = False
        msg = MsgBox("Valitettavasti ei ollut yhtään kappaletta.", 0, "Onnenpyörä")
        frmMain.tmrAika.Enabled = True
        frmTehtava.Visible = True
        frmKirjaimet.Visible = True
        frmMain.VaihdaVuoro
        frmMain.Palkinnonaika = 0
    End If

End If

If KSanottu = 1 Then
    frmMain.VaihdaVuoro
    frmMain.Palkinnonaika = 0
End If

If frmMain.YllatyksenVuoro = 0 Then
    lblPisteet.Caption = Format(KokoPistemaara)

    Select Case frmMain.Pelaaja

    Case 1
        frmMain.PisteetP1 = frmMain.PisteetP1 + KokoPistemaara

    Case 2
        frmMain.PisteetP2 = frmMain.PisteetP2 + KokoPistemaara

    Case 3
        frmMain.PisteetP3 = frmMain.PisteetP3 + KokoPistemaara

    End Select

    frmMain.Pisteet

End If

frmNapit.cmdPyoritys.Enabled = True
frmNapit.cmdOsta.Enabled = True
frmNapit.cmdRatkaise.Enabled = True
frmMain.Aika = 7
frmKirjaimet.tmrAika.Enabled = True
frmKirjaimet.Enabled = False
frmNapit.Enabled = True
frmNapit.SetFocus
KAika = 3
Yllatyspalkinto

End Sub

Private Sub Nayta_Ratkaisu()

Dim Vali As String

For i = 0 To 12
    Vali = Teksti(i)
    frmTehtava.txtKirjain(i).Text = Vali
Next

For i = 13 To 25
    Vali = Teksti2(i - 13)
    frmTehtava.txtKirjain(i).Text = Vali
Next

For i = 26 To 38
    Vali = Teksti3(i - 26)
    frmTehtava.txtKirjain(i).Text = Vali
Next

For i = 39 To 51
    Vali = Teksti4(i - 39)
    frmTehtava.txtKirjain(i).Text = Vali
Next

End Sub

Private Sub Tyhjaa_Ruudut()

For i = 0 To 51
    frmTehtava.txtKirjain(i).Text = ""
Next

End Sub

Private Sub ValitaanKirjaimet()

Valittujakirjaimia = 2
msg = MsgBox("Valitse yksi vokaali ja yksi konsonantti.", 0, "Onnenpyörä")
frmMain.VokaalitNakyviin
frmMain.KonsonantitNakyviin

End Sub

Public Sub Yllatyspalkinto()

If frmMain.Palkinnonaika <> 1 Then
    Exit Sub
Else

    frmTehtava.Visible = False
    frmKirjaimet.Visible = False
    frmMain.Aika = 7
    frmKirjaimet.tmrAika.Enabled = False
    frmMain.Palkinto = 2
    frmPalkinto.Show
    frmKirjaimet.tmrAika.Enabled = True

End If

End Sub

Private Sub cmdKonso_Click(Index As Integer)

If Bonus <> 1 Then
    Konso (Index)

Else
    BKonso (Index)

End If

End Sub

Private Sub Vokaali(Index As Integer)

Dim kirjain As String
Dim maara As Integer

maara = 0
KSanottu = 0

lblValitut(Index + 20).Visible = False
Select Case Index
    
Case 0
    If A = 0 Then
    A = 1
    kirjain = "A"
    For i = 0 To 12
        If Teksti(i) = "A" Then
            frmTehtava.txtKirjain(i).Text = "A"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "A" Then
            frmTehtava.txtKirjain(i).Text = "A"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "A" Then
            frmTehtava.txtKirjain(i).Text = "A"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "A" Then
            frmTehtava.txtKirjain(i).Text = "A"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If

Case 1
    If E = 0 Then
    E = 1
    kirjain = "E"
    For i = 0 To 12
        If Teksti(i) = "E" Then
            frmTehtava.txtKirjain(i).Text = "E"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "E" Then
            frmTehtava.txtKirjain(i).Text = "E"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "E" Then
            frmTehtava.txtKirjain(i).Text = "E"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "E" Then
            frmTehtava.txtKirjain(i).Text = "E"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If

Case 2
    If II = 0 Then
    II = 1
    kirjain = "I"
    For i = 0 To 12
        If Teksti(i) = "I" Then
            frmTehtava.txtKirjain(i).Text = "I"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "I" Then
            frmTehtava.txtKirjain(i).Text = "I"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "I" Then
            frmTehtava.txtKirjain(i).Text = "I"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "I" Then
            frmTehtava.txtKirjain(i).Text = "I"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If

Case 3
    If O = 0 Then
    O = 1
    kirjain = "O"
    For i = 0 To 12
        If Teksti(i) = "O" Then
            frmTehtava.txtKirjain(i).Text = "O"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "O" Then
            frmTehtava.txtKirjain(i).Text = "O"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "O" Then
            frmTehtava.txtKirjain(i).Text = "O"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "O" Then
            frmTehtava.txtKirjain(i).Text = "O"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If

Case 4
    If U = 0 Then
    U = 1
    kirjain = "U"
    For i = 0 To 12
        If Teksti(i) = "U" Then
            frmTehtava.txtKirjain(i).Text = "U"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "U" Then
            frmTehtava.txtKirjain(i).Text = "U"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "U" Then
            frmTehtava.txtKirjain(i).Text = "U"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "U" Then
            frmTehtava.txtKirjain(i).Text = "U"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If

Case 5
    If Y = 0 Then
    Y = 1
    kirjain = "Y"
    For i = 0 To 12
        If Teksti(i) = "Y" Then
            frmTehtava.txtKirjain(i).Text = "Y"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "Y" Then
            frmTehtava.txtKirjain(i).Text = "Y"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "Y" Then
            frmTehtava.txtKirjain(i).Text = "Y"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "Y" Then
            frmTehtava.txtKirjain(i).Text = "Y"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If

Case 6
    If AO = 0 Then
    AO = 1
    kirjain = "Å"
    For i = 0 To 12
        If Teksti(i) = "Å" Then
            frmTehtava.txtKirjain(i).Text = "Å"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "Å" Then
            frmTehtava.txtKirjain(i).Text = "Å"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "Å" Then
            frmTehtava.txtKirjain(i).Text = "Å"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "Å" Then
            frmTehtava.txtKirjain(i).Text = "Å"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If

Case 7
    If AE = 0 Then
    AE = 1
    kirjain = "Ä"
    For i = 0 To 12
        If Teksti(i) = "Ä" Then
            frmTehtava.txtKirjain(i).Text = "Ä"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "Ä" Then
            frmTehtava.txtKirjain(i).Text = "Ä"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "Ä" Then
            frmTehtava.txtKirjain(i).Text = "Ä"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "Ä" Then
            frmTehtava.txtKirjain(i).Text = "Ä"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If

Case 8
    If OE = 0 Then
    OE = 1
    kirjain = "Ö"
    For i = 0 To 12
        If Teksti(i) = "Ö" Then
            frmTehtava.txtKirjain(i).Text = "Ö"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "Ö" Then
            frmTehtava.txtKirjain(i).Text = "Ö"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "Ö" Then
            frmTehtava.txtKirjain(i).Text = "Ö"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "Ö" Then
            frmTehtava.txtKirjain(i).Text = "Ö"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            maara = maara + 1
        End If
    Next
    Else
        Kirjainsanottu
    End If

End Select

If frmMain.YllatyksenVuoro = 0 And maara <> 0 Then
    Select Case frmMain.Pelaaja
    Case 1
        frmMain.PisteetP1 = frmMain.PisteetP1 - 300
        frmMain.Pisteet
        lblPisteet.Caption = "-300"
    Case 2
        frmMain.PisteetP2 = frmMain.PisteetP2 - 300
        frmMain.Pisteet
        lblPisteet.Caption = "-300"
    Case 3
        frmMain.PisteetP3 = frmMain.PisteetP3 - 300
        frmMain.Pisteet
        lblPisteet.Caption = "-300"
    End Select
End If

If frmMain.YllatyksenVuoro = 0 And maara = 0 Then
    lblPisteet.Caption = ""
End If

lblMaara.Caption = maara
lblKirjain.Caption = kirjain
frmMain.KonsonantitPiiloon
frmMain.VokaalitPiiloon

If KSanottu <> 1 Then
    
    If maara = 0 Then
        frmTehtava.Visible = False
        frmKirjaimet.Visible = False
        frmMain.tmrAika.Enabled = False
        msg = MsgBox("Valitettavasti ei ollut yhtään kappaletta.", 0, "Onnenpyörä")
        frmMain.tmrAika.Enabled = True
        frmTehtava.Visible = True
        frmKirjaimet.Visible = True
        frmMain.VaihdaVuoro
        frmMain.Palkinnonaika = 0
    End If

End If

If KSanottu = 1 Then
    frmMain.VaihdaVuoro
    frmMain.Palkinnonaika = 0
End If

frmNapit.cmdPyoritys.Enabled = True
frmMain.Aika = 7
frmNapit.cmdRatkaise.Enabled = True
frmKirjaimet.tmrAika.Enabled = True
frmKirjaimet.Enabled = False
frmNapit.Enabled = True
frmNapit.SetFocus
KAika = 3
Yllatyspalkinto

End Sub
Private Sub BVokaali(Index As Integer)

Dim kirjain As String

BSanottu = 0
lblValitut(Index + 20).Visible = False

Select Case Index
    
Case 0
    If A = 0 Then
        A = 1
        kirjain = "A"
        
        For i = 0 To 12
            
            If Teksti(i) = "A" Then
                frmTehtava.txtKirjain(i).Text = "A"
                frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
                
            End If
        
        Next
        
        For i = 13 To 25
            
            If Teksti2(i - 13) = "A" Then
                frmTehtava.txtKirjain(i).Text = "A"
                frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
                
            End If
        Next
        
        For i = 26 To 38
            If Teksti3(i - 26) = "A" Then
                frmTehtava.txtKirjain(i).Text = "A"
                frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
                
            End If
        Next
        
        For i = 39 To 51
            If Teksti4(i - 39) = "A" Then
                frmTehtava.txtKirjain(i).Text = "A"
                frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
                
            End If
        Next
    Else
        
        BKirjainsanottu
    
    End If

Case 1
    If E = 0 Then
    E = 1
    kirjain = "E"
    For i = 0 To 12
        If Teksti(i) = "E" Then
            frmTehtava.txtKirjain(i).Text = "E"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "E" Then
            frmTehtava.txtKirjain(i).Text = "E"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "E" Then
            frmTehtava.txtKirjain(i).Text = "E"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "E" Then
            frmTehtava.txtKirjain(i).Text = "E"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    Else
        BKirjainsanottu
    End If

Case 2
    If II = 0 Then
    II = 1
    kirjain = "I"
    For i = 0 To 12
        If Teksti(i) = "I" Then
            frmTehtava.txtKirjain(i).Text = "I"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "I" Then
            frmTehtava.txtKirjain(i).Text = "I"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "I" Then
            frmTehtava.txtKirjain(i).Text = "I"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "I" Then
            frmTehtava.txtKirjain(i).Text = "I"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    Else
        BKirjainsanottu
    End If

Case 3
    If O = 0 Then
    O = 1
    kirjain = "O"
    For i = 0 To 12
        If Teksti(i) = "O" Then
            frmTehtava.txtKirjain(i).Text = "O"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "O" Then
            frmTehtava.txtKirjain(i).Text = "O"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "O" Then
            frmTehtava.txtKirjain(i).Text = "O"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "O" Then
            frmTehtava.txtKirjain(i).Text = "O"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    Else
        BKirjainsanottu
    End If

Case 4
    If U = 0 Then
    U = 1
    kirjain = "U"
    For i = 0 To 12
        If Teksti(i) = "U" Then
            frmTehtava.txtKirjain(i).Text = "U"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "U" Then
            frmTehtava.txtKirjain(i).Text = "U"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "U" Then
            frmTehtava.txtKirjain(i).Text = "U"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "U" Then
            frmTehtava.txtKirjain(i).Text = "U"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    Else
        BKirjainsanottu
    End If

Case 5
    If Y = 0 Then
    Y = 1
    kirjain = "Y"
    For i = 0 To 12
        If Teksti(i) = "Y" Then
            frmTehtava.txtKirjain(i).Text = "Y"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "Y" Then
            frmTehtava.txtKirjain(i).Text = "Y"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "Y" Then
            frmTehtava.txtKirjain(i).Text = "Y"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "Y" Then
            frmTehtava.txtKirjain(i).Text = "Y"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    Else
        BKirjainsanottu
    End If

Case 6
    If AO = 0 Then
    AO = 1
    kirjain = "Å"
    For i = 0 To 12
        If Teksti(i) = "Å" Then
            frmTehtava.txtKirjain(i).Text = "Å"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "Å" Then
            frmTehtava.txtKirjain(i).Text = "Å"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "Å" Then
            frmTehtava.txtKirjain(i).Text = "Å"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "Å" Then
            frmTehtava.txtKirjain(i).Text = "Å"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    Else
        BKirjainsanottu
    End If

Case 7
    If AE = 0 Then
    AE = 1
    kirjain = "Ä"
    For i = 0 To 12
        If Teksti(i) = "Ä" Then
            frmTehtava.txtKirjain(i).Text = "Ä"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "Ä" Then
            frmTehtava.txtKirjain(i).Text = "Ä"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "Ä" Then
            frmTehtava.txtKirjain(i).Text = "Ä"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "Ä" Then
            frmTehtava.txtKirjain(i).Text = "Ä"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    Else
        BKirjainsanottu
    End If

Case 8
    If OE = 0 Then
    OE = 1
    kirjain = "Ö"
    For i = 0 To 12
        If Teksti(i) = "Ö" Then
            frmTehtava.txtKirjain(i).Text = "Ö"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 13 To 25
        If Teksti2(i - 13) = "Ö" Then
            frmTehtava.txtKirjain(i).Text = "Ö"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 26 To 38
        If Teksti3(i - 26) = "Ö" Then
            frmTehtava.txtKirjain(i).Text = "Ö"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    For i = 39 To 51
        If Teksti4(i - 39) = "Ö" Then
            frmTehtava.txtKirjain(i).Text = "Ö"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
            
        End If
    Next
    Else
        BKirjainsanottu
    End If

End Select

If BSanottu = 0 Then
    Valittujakirjaimia = Valittujakirjaimia - 1
    
    Select Case Valittujakirjaimia
    
    Case 1
        frmMain.VokaalitPiiloon
    
    Case 0
        frmMain.VokaalitPiiloon
        BonusAikaAlkaa
    
    End Select

End If

End Sub

Public Sub Aloitus(Index As Integer)

tehtava = Int(Index)
Avaa_Tehtava (tehtava)
Tekstit_Muistiin
Tekstit_Naytolle
Vaihda_Varit_Alku
Aakkosiin_kuulumattomat

End Sub

Public Sub BonusAloitus(Index As Integer)

Valiaika = 2
tehtava = Int(Index)
Avaa_BonusTehtava (tehtava)
Tekstit_Muistiin
Tekstit_Naytolle
Vaihda_Varit_Alku
Aakkosiin_kuulumattomat

End Sub

Private Sub Tekstit_Muistiin()

k1 = Mid(Tehtavarivi, 5, 1)
k2 = Mid(Tehtavarivi, 6, 1)
k3 = Mid(Tehtavarivi, 7, 1)
k4 = Mid(Tehtavarivi, 8, 1)
k5 = Mid(Tehtavarivi, 9, 1)
k6 = Mid(Tehtavarivi, 10, 1)
k7 = Mid(Tehtavarivi, 11, 1)
k8 = Mid(Tehtavarivi, 12, 1)
k9 = Mid(Tehtavarivi, 13, 1)
k10 = Mid(Tehtavarivi, 14, 1)
k11 = Mid(Tehtavarivi, 15, 1)
k12 = Mid(Tehtavarivi, 16, 1)
k13 = Mid(Tehtavarivi, 17, 1)

Teksti = Array(k1, k2, k3, k4, k5, k6, k7, k8, k9, k10, k11, k12, k13)

k1 = Mid(Tehtavarivi, 19, 1)
k2 = Mid(Tehtavarivi, 20, 1)
k3 = Mid(Tehtavarivi, 21, 1)
k4 = Mid(Tehtavarivi, 22, 1)
k5 = Mid(Tehtavarivi, 23, 1)
k6 = Mid(Tehtavarivi, 24, 1)
k7 = Mid(Tehtavarivi, 25, 1)
k8 = Mid(Tehtavarivi, 26, 1)
k9 = Mid(Tehtavarivi, 27, 1)
k10 = Mid(Tehtavarivi, 28, 1)
k11 = Mid(Tehtavarivi, 29, 1)
k12 = Mid(Tehtavarivi, 30, 1)
k13 = Mid(Tehtavarivi, 31, 1)

Teksti2 = Array(k1, k2, k3, k4, k5, k6, k7, k8, k9, k10, k11, k12, k13)

k1 = Mid(Tehtavarivi, 33, 1)
k2 = Mid(Tehtavarivi, 34, 1)
k3 = Mid(Tehtavarivi, 35, 1)
k4 = Mid(Tehtavarivi, 36, 1)
k5 = Mid(Tehtavarivi, 37, 1)
k6 = Mid(Tehtavarivi, 38, 1)
k7 = Mid(Tehtavarivi, 39, 1)
k8 = Mid(Tehtavarivi, 40, 1)
k9 = Mid(Tehtavarivi, 41, 1)
k10 = Mid(Tehtavarivi, 42, 1)
k11 = Mid(Tehtavarivi, 43, 1)
k12 = Mid(Tehtavarivi, 44, 1)
k13 = Mid(Tehtavarivi, 45, 1)

Teksti3 = Array(k1, k2, k3, k4, k5, k6, k7, k8, k9, k10, k11, k12, k13)

k1 = Mid(Tehtavarivi, 47, 1)
k2 = Mid(Tehtavarivi, 48, 1)
k3 = Mid(Tehtavarivi, 49, 1)
k4 = Mid(Tehtavarivi, 50, 1)
k5 = Mid(Tehtavarivi, 51, 1)
k6 = Mid(Tehtavarivi, 52, 1)
k7 = Mid(Tehtavarivi, 53, 1)
k8 = Mid(Tehtavarivi, 54, 1)
k9 = Mid(Tehtavarivi, 55, 1)
k10 = Mid(Tehtavarivi, 56, 1)
k11 = Mid(Tehtavarivi, 57, 1)
k12 = Mid(Tehtavarivi, 58, 1)
k13 = Mid(Tehtavarivi, 59, 1)

Teksti4 = Array(k1, k2, k3, k4, k5, k6, k7, k8, k9, k10, k11, k12, k13)

End Sub


Private Sub Tekstit_Naytolle()

Dim Vali As String

For i = 0 To 12
    Vali = Teksti(i)
    
    If Asc(Vali) = 32 Then
        frmTehtava.txtKirjain(i).Text = Vali
    
    End If

Next

For i = 13 To 25
    Vali = Teksti2(i - 13)
    
    If Asc(Vali) = 32 Then
        frmTehtava.txtKirjain(i).Text = Vali
    
    End If

Next

For i = 26 To 38
    Vali = Teksti3(i - 26)
    
    If Asc(Vali) = 32 Then
        frmTehtava.txtKirjain(i).Text = Vali
    
    End If

Next

For i = 39 To 51
    Vali = Teksti4(i - 39)
    
    If Asc(Vali) = 32 Then
        frmTehtava.txtKirjain(i).Text = Vali
    
    End If

Next

End Sub


Private Sub Vaihda_Varit_Alku()

frmTehtava.txtAihe.Text = Mid(Tehtavarivi, 61, 17)

For i = 0 To 51
    frmTehtava.txtKirjain(i).BackColor = &HC0C0C0
    
    If frmTehtava.txtKirjain(i).Text = " " Then
        frmTehtava.txtKirjain(i).BackColor = &HFF0000
    
    End If

Next

End Sub

Private Sub Vaihda_Varit_Ratkaisu()

For i = 0 To 51
    frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
    
    If frmTehtava.txtKirjain(i).Text = " " Then
        frmTehtava.txtKirjain(i).BackColor = &HFF0000
    
    End If

Next

End Sub




Public Sub Ratkaise()

Nayta_Ratkaisu
Vaihda_Varit_Ratkaisu
frmMain.Ratkaistu

End Sub

Public Sub BonusRatkaise()

Nayta_Ratkaisu
Vaihda_Varit_Ratkaisu
BonusRatkaistu

End Sub

Public Sub BonusEiRatkaise()

Dim msg
Dim oikea As String
oikea = "Oikea ratkaisu olisi ollut: " + Tehtavateksti
msg = MsgBox(oikea, 0, "Onnenpyörä")

Nayta_Ratkaisu
Vaihda_Varit_Ratkaisu
BonusEiRatkaistu

End Sub

Private Sub Aakkosiin_kuulumattomat()

Dim val As String

For i = 0 To 12
    val = Teksti(i)
    
    If val <> "" Then
        
        Select Case Asc(val)
        
        Case 33
            frmTehtava.txtKirjain(i).Text = "!"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 34
            frmTehtava.txtKirjain(i).Text = Chr(34)
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 35
            frmTehtava.txtKirjain(i).Text = "#"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 37
            frmTehtava.txtKirjain(i).Text = "%"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 38
            frmTehtava.txtKirjain(i).Text = "&"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 39
            frmTehtava.txtKirjain(i).Text = "'"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 44
            frmTehtava.txtKirjain(i).Text = ","
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 45
            frmTehtava.txtKirjain(i).Text = "-"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 46
            frmTehtava.txtKirjain(i).Text = "."
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 47
            frmTehtava.txtKirjain(i).Text = "/"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 48
            frmTehtava.txtKirjain(i).Text = "0"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 49
            frmTehtava.txtKirjain(i).Text = "1"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 50
            frmTehtava.txtKirjain(i).Text = "2"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 51
            frmTehtava.txtKirjain(i).Text = "3"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 52
            frmTehtava.txtKirjain(i).Text = "4"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 53
            frmTehtava.txtKirjain(i).Text = "5"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 54
            frmTehtava.txtKirjain(i).Text = "6"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 55
            frmTehtava.txtKirjain(i).Text = "7"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 56
            frmTehtava.txtKirjain(i).Text = "8"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 57
            frmTehtava.txtKirjain(i).Text = "9"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 58
            frmTehtava.txtKirjain(i).Text = ":"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 59
            frmTehtava.txtKirjain(i).Text = ";"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 61
            frmTehtava.txtKirjain(i).Text = "="
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        End Select
    
    End If

Next

For i = 13 To 25
    val = Teksti2(i - 13)
    
    If val <> "" Then
        
        Select Case Asc(val)
        
        Case 33
            frmTehtava.txtKirjain(i).Text = "!"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 34
            frmTehtava.txtKirjain(i).Text = Chr(34)
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 35
            frmTehtava.txtKirjain(i).Text = "#"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 37
            frmTehtava.txtKirjain(i).Text = "%"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 38
            frmTehtava.txtKirjain(i).Text = "&"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 39
            frmTehtava.txtKirjain(i).Text = "'"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 44
            frmTehtava.txtKirjain(i).Text = ","
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 45
            frmTehtava.txtKirjain(i).Text = "-"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 46
            frmTehtava.txtKirjain(i).Text = "."
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 47
            frmTehtava.txtKirjain(i).Text = "/"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 48
            frmTehtava.txtKirjain(i).Text = "0"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 49
            frmTehtava.txtKirjain(i).Text = "1"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 50
            frmTehtava.txtKirjain(i).Text = "2"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 51
            frmTehtava.txtKirjain(i).Text = "3"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 52
            frmTehtava.txtKirjain(i).Text = "4"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 53
            frmTehtava.txtKirjain(i).Text = "5"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 54
            frmTehtava.txtKirjain(i).Text = "6"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 55
            frmTehtava.txtKirjain(i).Text = "7"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 56
            frmTehtava.txtKirjain(i).Text = "8"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 57
            frmTehtava.txtKirjain(i).Text = "9"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 58
            frmTehtava.txtKirjain(i).Text = ":"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 59
            frmTehtava.txtKirjain(i).Text = ";"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 61
            frmTehtava.txtKirjain(i).Text = "="
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        End Select
    
    End If

Next

For i = 26 To 38
    val = Teksti3(i - 26)
    
    If val <> "" Then
        
        Select Case Asc(val)
        
        Case 33
            frmTehtava.txtKirjain(i).Text = "!"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 34
            frmTehtava.txtKirjain(i).Text = Chr(34)
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 35
            frmTehtava.txtKirjain(i).Text = "#"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 37
            frmTehtava.txtKirjain(i).Text = "%"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 38
            frmTehtava.txtKirjain(i).Text = "&"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 39
            frmTehtava.txtKirjain(i).Text = "'"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 44
            frmTehtava.txtKirjain(i).Text = ","
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 45
            frmTehtava.txtKirjain(i).Text = "-"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 46
            frmTehtava.txtKirjain(i).Text = "."
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 47
            frmTehtava.txtKirjain(i).Text = "/"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 48
            frmTehtava.txtKirjain(i).Text = "0"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 49
            frmTehtava.txtKirjain(i).Text = "1"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 50
            frmTehtava.txtKirjain(i).Text = "2"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 51
            frmTehtava.txtKirjain(i).Text = "3"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 52
            frmTehtava.txtKirjain(i).Text = "4"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 53
            frmTehtava.txtKirjain(i).Text = "5"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 54
            frmTehtava.txtKirjain(i).Text = "6"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 55
            frmTehtava.txtKirjain(i).Text = "7"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 56
            frmTehtava.txtKirjain(i).Text = "8"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 57
            frmTehtava.txtKirjain(i).Text = "9"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 58
            frmTehtava.txtKirjain(i).Text = ":"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 59
            frmTehtava.txtKirjain(i).Text = ";"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 61
            frmTehtava.txtKirjain(i).Text = "="
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        End Select
    
    End If

Next

For i = 39 To 51
    val = Teksti4(i - 39)
    
    If val <> "" Then
        
        Select Case Asc(val)
        
        Case 33
            frmTehtava.txtKirjain(i).Text = "!"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 34
            frmTehtava.txtKirjain(i).Text = Chr(34)
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 35
            frmTehtava.txtKirjain(i).Text = "#"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 37
            frmTehtava.txtKirjain(i).Text = "%"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 38
            frmTehtava.txtKirjain(i).Text = "&"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 39
            frmTehtava.txtKirjain(i).Text = "'"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 44
            frmTehtava.txtKirjain(i).Text = ","
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 45
            frmTehtava.txtKirjain(i).Text = "-"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 46
            frmTehtava.txtKirjain(i).Text = "."
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 47
            frmTehtava.txtKirjain(i).Text = "/"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 48
            frmTehtava.txtKirjain(i).Text = "0"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 49
            frmTehtava.txtKirjain(i).Text = "1"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 50
            frmTehtava.txtKirjain(i).Text = "2"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 51
            frmTehtava.txtKirjain(i).Text = "3"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 52
            frmTehtava.txtKirjain(i).Text = "4"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 53
            frmTehtava.txtKirjain(i).Text = "5"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 54
            frmTehtava.txtKirjain(i).Text = "6"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 55
            frmTehtava.txtKirjain(i).Text = "7"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 56
            frmTehtava.txtKirjain(i).Text = "8"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 57
            frmTehtava.txtKirjain(i).Text = "9"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 58
            frmTehtava.txtKirjain(i).Text = ":"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 59
            frmTehtava.txtKirjain(i).Text = ";"
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        Case 61
            frmTehtava.txtKirjain(i).Text = "="
            frmTehtava.txtKirjain(i).BackColor = &HFFFFFF
        
        End Select
    
    End If

Next

End Sub

Private Sub cmdVokaali_Click(Index As Integer)

If Bonus <> 1 Then
    Vokaali (Index)
Else
    BVokaali (Index)
End If

End Sub

Private Sub Form_Load()

Select Case Screen.Width / 15
Case 640
    frmKirjaimet.Top = 3000
    frmKirjaimet.Left = 4700
Case 800
    frmKirjaimet.Top = 4400
    frmKirjaimet.Left = 6500
Case Else
    frmKirjaimet.Top = 6500
    frmKirjaimet.Left = 9000
End Select

End Sub

Private Sub tmrAika_Timer()

KAika = KAika - 1

If KAika = 0 Then
    lblMaara.Caption = ""
    lblPisteet.Caption = ""
    lblKirjain.Caption = ""
    tmrAika.Enabled = False

End If

End Sub


Private Sub tmrAika2_Timer()

Valiaika = Valiaika - 1

If Valiaika = 0 Then
    BonusAloitus2

End If

End Sub
