VERSION 5.00
Begin VB.Form frmTehtava 
   BackColor       =   &H001F15D9&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Onnenpyörä - Tehtävä"
   ClientHeight    =   2595
   ClientLeft      =   7785
   ClientTop       =   960
   ClientWidth     =   6330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2595
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtAihe 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   52
      Top             =   2040
      Width           =   5175
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   51
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   51
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   50
      Left            =   5400
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   50
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   49
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   49
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   48
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   48
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   47
      Left            =   3960
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   47
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   46
      Left            =   3480
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   46
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   45
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   45
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   44
      Left            =   2520
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   44
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   43
      Left            =   2040
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   43
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   42
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   42
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   41
      Left            =   1080
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   41
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   40
      Left            =   600
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   40
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   39
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   39
      Top             =   1560
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   38
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   38
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   37
      Left            =   5400
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   37
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   36
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   36
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   35
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   35
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   34
      Left            =   3960
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   34
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   33
      Left            =   3480
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   33
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   32
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   32
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   31
      Left            =   2520
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   31
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   30
      Left            =   2040
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   30
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   29
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   29
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   28
      Left            =   1080
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   28
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   27
      Left            =   600
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   27
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   26
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   26
      Top             =   1080
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   25
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   25
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   24
      Left            =   5400
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   24
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   23
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   23
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   22
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   22
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   21
      Left            =   3960
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   21
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   20
      Left            =   3480
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   20
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   19
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   19
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   18
      Left            =   2520
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   18
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   17
      Left            =   2040
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   17
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   16
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   16
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   15
      Left            =   1080
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   15
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   14
      Left            =   600
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   14
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   13
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   13
      Top             =   600
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   12
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   12
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   11
      Left            =   5400
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   11
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   10
      Left            =   4920
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   9
      Left            =   4440
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   9
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   8
      Left            =   3960
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   8
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   7
      Left            =   3480
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   7
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   6
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   5
      Left            =   2520
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   5
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   4
      Left            =   2040
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   4
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   3
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   1080
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   2
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   600
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtKirjain 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmTehtava"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Select Case Screen.Width / 15
Case 640
    frmTehtava.Top = 0
    frmTehtava.Left = 3000
Case 800
    frmTehtava.Top = 200
    frmTehtava.Left = 4950
Case Else
    frmTehtava.Top = 675
    frmTehtava.Left = 7750
End Select

End Sub


