VERSION 5.00
Object = "{4EC837EB-4201-4C6D-A7D7-0C99D69CD9A6}#1.0#0"; "GradientCommand.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmverreportes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ver Reportes"
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   14835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7395
   ScaleWidth      =   14835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   13150
      BackColor       =   15649666
      Caption         =   "Bordados Marion"
      CaptionEstilo3D =   1
      BackColor       =   15649666
      ForeColor       =   16711680
      ColorBarraArriba=   15649666
      ColorBarraAbajo =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorTextShadow =   16777215
      Begin GradientCommand.GGCommand cmdanterior 
         Height          =   615
         Left            =   1680
         TabIndex        =   3
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
         Caption         =   "Anterior"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientToColor =   15649666
      End
      Begin GradientCommand.GGCommand cmdsiguiente 
         Height          =   615
         Left            =   4800
         TabIndex        =   5
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
         Caption         =   "Siguiente"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientToColor =   15649666
      End
      Begin GradientCommand.GGCommand cmdactual 
         Height          =   615
         Left            =   3240
         TabIndex        =   4
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   1085
         Caption         =   "Actual"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientToColor =   15649666
      End
      Begin GradientCommand.GGCommand cmdgenerar 
         Height          =   495
         Left            =   6360
         TabIndex        =   9
         Top             =   2280
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   873
         Caption         =   "Generar"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GradientToColor =   15649666
      End
      Begin TabDlg.SSTab sstreporte 
         Height          =   3975
         Left            =   120
         TabIndex        =   12
         Top             =   3240
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   7011
         _Version        =   393216
         TabHeight       =   520
         BackColor       =   15649666
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Semana"
         TabPicture(0)   =   "frmverreportes.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "mscreporte"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "General"
         TabPicture(1)   =   "frmverreportes.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "mscreportegen"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Rendimiento"
         TabPicture(2)   =   "frmverreportes.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label5"
         Tab(2).Control(1)=   "Label6"
         Tab(2).Control(2)=   "Label9"
         Tab(2).Control(3)=   "Label7(0)"
         Tab(2).Control(4)=   "Label8"
         Tab(2).Control(5)=   "Label10"
         Tab(2).Control(6)=   "Label12"
         Tab(2).Control(7)=   "Label13"
         Tab(2).Control(8)=   "Label14"
         Tab(2).Control(9)=   "Label15"
         Tab(2).Control(10)=   "Label16"
         Tab(2).Control(11)=   "Label17"
         Tab(2).Control(12)=   "Label18"
         Tab(2).Control(13)=   "Label19"
         Tab(2).Control(14)=   "lbloperario"
         Tab(2).Control(15)=   "Label7(1)"
         Tab(2).Control(16)=   "Label20"
         Tab(2).Control(17)=   "Label21"
         Tab(2).Control(18)=   "Label22"
         Tab(2).Control(19)=   "Label7(2)"
         Tab(2).Control(20)=   "Label23"
         Tab(2).Control(21)=   "Label24"
         Tab(2).Control(22)=   "Label25"
         Tab(2).Control(23)=   "Label7(4)"
         Tab(2).Control(24)=   "Label29"
         Tab(2).Control(25)=   "Label30"
         Tab(2).Control(26)=   "Label31"
         Tab(2).Control(27)=   "Label7(5)"
         Tab(2).Control(28)=   "Label32"
         Tab(2).Control(29)=   "Label33"
         Tab(2).Control(30)=   "Label34"
         Tab(2).Control(31)=   "Label7(6)"
         Tab(2).Control(32)=   "Label35"
         Tab(2).Control(33)=   "Label36"
         Tab(2).Control(34)=   "Label37"
         Tab(2).Control(35)=   "Label7(7)"
         Tab(2).Control(36)=   "Label38"
         Tab(2).Control(37)=   "Label39"
         Tab(2).Control(38)=   "Label40"
         Tab(2).Control(39)=   "txtsemanade"
         Tab(2).Control(40)=   "txtsemanaa"
         Tab(2).Control(41)=   "txtesperadaslu"
         Tab(2).Control(42)=   "txtestimadaslu"
         Tab(2).Control(43)=   "txtrealeslu"
         Tab(2).Control(44)=   "txteficiencialu"
         Tab(2).Control(45)=   "txteficienciama"
         Tab(2).Control(46)=   "txtrealesma"
         Tab(2).Control(47)=   "txtestimadasma"
         Tab(2).Control(48)=   "txtesperadasma"
         Tab(2).Control(49)=   "txteficienciami"
         Tab(2).Control(50)=   "txtrealesmi"
         Tab(2).Control(51)=   "txtestimadasmi"
         Tab(2).Control(52)=   "txtesperadasmi"
         Tab(2).Control(53)=   "txteficienciaju"
         Tab(2).Control(54)=   "txtrealesju"
         Tab(2).Control(55)=   "txtestimadasju"
         Tab(2).Control(56)=   "txtesperadasju"
         Tab(2).Control(57)=   "txteficienciavi"
         Tab(2).Control(58)=   "txtrealesvi"
         Tab(2).Control(59)=   "txtestimadasvi"
         Tab(2).Control(60)=   "txtesperadasvi"
         Tab(2).Control(61)=   "txteficienciasa"
         Tab(2).Control(62)=   "txtrealessa"
         Tab(2).Control(63)=   "txtestimadassa"
         Tab(2).Control(64)=   "txtesperadassa"
         Tab(2).Control(65)=   "txteficienciado"
         Tab(2).Control(66)=   "txtrealesdo"
         Tab(2).Control(67)=   "txtestimadasdo"
         Tab(2).Control(68)=   "txtesperadasdo"
         Tab(2).ControlCount=   69
         Begin MSChart20Lib.MSChart mscreporte 
            Height          =   3495
            Left            =   240
            OleObjectBlob   =   "frmverreportes.frx":0054
            TabIndex        =   13
            Top             =   360
            Width           =   14175
         End
         Begin VB.TextBox txtesperadasdo 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -61440
            TabIndex        =   87
            Top             =   1980
            Width           =   855
         End
         Begin VB.TextBox txtestimadasdo 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -61440
            TabIndex        =   86
            Top             =   2460
            Width           =   855
         End
         Begin VB.TextBox txtrealesdo 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -61440
            TabIndex        =   85
            Top             =   2940
            Width           =   855
         End
         Begin VB.TextBox txteficienciado 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -61440
            TabIndex        =   84
            Top             =   3420
            Width           =   855
         End
         Begin VB.TextBox txtesperadassa 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -63480
            TabIndex        =   83
            Top             =   1980
            Width           =   855
         End
         Begin VB.TextBox txtestimadassa 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -63480
            TabIndex        =   82
            Top             =   2460
            Width           =   855
         End
         Begin VB.TextBox txtrealessa 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -63480
            TabIndex        =   81
            Top             =   2940
            Width           =   855
         End
         Begin VB.TextBox txteficienciasa 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -63480
            TabIndex        =   80
            Top             =   3420
            Width           =   855
         End
         Begin VB.TextBox txtesperadasvi 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -65520
            TabIndex        =   79
            Top             =   1980
            Width           =   855
         End
         Begin VB.TextBox txtestimadasvi 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -65520
            TabIndex        =   78
            Top             =   2460
            Width           =   855
         End
         Begin VB.TextBox txtrealesvi 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -65520
            TabIndex        =   77
            Top             =   2940
            Width           =   855
         End
         Begin VB.TextBox txteficienciavi 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -65520
            TabIndex        =   76
            Top             =   3420
            Width           =   855
         End
         Begin VB.TextBox txtesperadasju 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -67560
            TabIndex        =   75
            Top             =   1980
            Width           =   855
         End
         Begin VB.TextBox txtestimadasju 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -67560
            TabIndex        =   74
            Top             =   2460
            Width           =   855
         End
         Begin VB.TextBox txtrealesju 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -67560
            TabIndex        =   73
            Top             =   2940
            Width           =   855
         End
         Begin VB.TextBox txteficienciaju 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -67560
            TabIndex        =   72
            Top             =   3420
            Width           =   855
         End
         Begin VB.TextBox txtesperadasmi 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -69600
            TabIndex        =   71
            Top             =   1980
            Width           =   855
         End
         Begin VB.TextBox txtestimadasmi 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -69600
            TabIndex        =   70
            Top             =   2460
            Width           =   855
         End
         Begin VB.TextBox txtrealesmi 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -69600
            TabIndex        =   69
            Top             =   2940
            Width           =   855
         End
         Begin VB.TextBox txteficienciami 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -69600
            TabIndex        =   68
            Top             =   3420
            Width           =   855
         End
         Begin VB.TextBox txtesperadasma 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -71640
            TabIndex        =   67
            Top             =   1980
            Width           =   855
         End
         Begin VB.TextBox txtestimadasma 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -71640
            TabIndex        =   66
            Top             =   2460
            Width           =   855
         End
         Begin VB.TextBox txtrealesma 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -71640
            TabIndex        =   65
            Top             =   2940
            Width           =   855
         End
         Begin VB.TextBox txteficienciama 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -71640
            TabIndex        =   64
            Top             =   3420
            Width           =   855
         End
         Begin VB.TextBox txteficiencialu 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -73680
            TabIndex        =   27
            Top             =   3420
            Width           =   855
         End
         Begin VB.TextBox txtrealeslu 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -73680
            TabIndex        =   25
            Top             =   2940
            Width           =   855
         End
         Begin VB.TextBox txtestimadaslu 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -73680
            TabIndex        =   23
            Top             =   2460
            Width           =   855
         End
         Begin VB.TextBox txtesperadaslu 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -73680
            TabIndex        =   21
            Top             =   1980
            Width           =   855
         End
         Begin VB.TextBox txtsemanaa 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -72240
            TabIndex        =   19
            Top             =   1140
            Width           =   1215
         End
         Begin VB.TextBox txtsemanade 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -73920
            TabIndex        =   18
            Top             =   1140
            Width           =   1215
         End
         Begin MSChart20Lib.MSChart mscreportegen 
            Height          =   3495
            Left            =   -74760
            OleObjectBlob   =   "frmverreportes.frx":2316
            TabIndex        =   14
            Top             =   360
            Width           =   14175
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            Caption         =   "Eficiencia:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -62520
            TabIndex        =   63
            Top             =   3420
            Width           =   1020
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            Caption         =   "Reales:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -62520
            TabIndex        =   62
            Top             =   2940
            Width           =   750
         End
         Begin VB.Label Label38 
            AutoSize        =   -1  'True
            Caption         =   "Estimadas:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -62520
            TabIndex        =   61
            Top             =   2460
            Width           =   1080
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Esperadas:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   -62520
            TabIndex        =   60
            Top             =   1980
            Width           =   1095
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "Eficiencia:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -64560
            TabIndex        =   59
            Top             =   3420
            Width           =   1020
         End
         Begin VB.Label Label36 
            AutoSize        =   -1  'True
            Caption         =   "Reales:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -64560
            TabIndex        =   58
            Top             =   2940
            Width           =   750
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Estimadas:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -64560
            TabIndex        =   57
            Top             =   2460
            Width           =   1080
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Esperadas:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   -64560
            TabIndex        =   56
            Top             =   1980
            Width           =   1095
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Eficiencia:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -66600
            TabIndex        =   55
            Top             =   3420
            Width           =   1020
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Reales:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -66600
            TabIndex        =   54
            Top             =   2940
            Width           =   750
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "Estimadas:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -66600
            TabIndex        =   53
            Top             =   2460
            Width           =   1080
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Esperadas:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   -66600
            TabIndex        =   52
            Top             =   1980
            Width           =   1095
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Eficiencia:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -68640
            TabIndex        =   51
            Top             =   3420
            Width           =   1020
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Reales:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -68640
            TabIndex        =   50
            Top             =   2940
            Width           =   750
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Estimadas:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -68640
            TabIndex        =   49
            Top             =   2460
            Width           =   1080
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Esperadas:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   -68640
            TabIndex        =   48
            Top             =   1980
            Width           =   1095
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Eficiencia:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -70680
            TabIndex        =   43
            Top             =   3420
            Width           =   1020
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Reales:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -70680
            TabIndex        =   42
            Top             =   2940
            Width           =   750
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Estimadas:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -70680
            TabIndex        =   41
            Top             =   2460
            Width           =   1080
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Esperadas:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   -70680
            TabIndex        =   40
            Top             =   1980
            Width           =   1095
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Eficiencia:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -72720
            TabIndex        =   39
            Top             =   3420
            Width           =   1020
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Reales:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -72720
            TabIndex        =   38
            Top             =   2940
            Width           =   750
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Estimadas:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -72720
            TabIndex        =   37
            Top             =   2460
            Width           =   1080
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Esperadas:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   -72720
            TabIndex        =   36
            Top             =   1980
            Width           =   1095
         End
         Begin VB.Label lbloperario 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -73680
            TabIndex        =   35
            Top             =   780
            Width           =   60
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Domingo"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   -62400
            TabIndex        =   34
            Top             =   1620
            Width           =   885
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Sabado"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   -64440
            TabIndex        =   33
            Top             =   1620
            Width           =   690
         End
         Begin VB.Label Label17 
            AutoSize        =   -1  'True
            Caption         =   "Viernes"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   -66480
            TabIndex        =   32
            Top             =   1620
            Width           =   765
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Jueves"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   -68520
            TabIndex        =   31
            Top             =   1620
            Width           =   705
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "Miercoles"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   -70680
            TabIndex        =   30
            Top             =   1620
            Width           =   1005
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Martes"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   -72600
            TabIndex        =   29
            Top             =   1620
            Width           =   720
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Lunes"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   -74640
            TabIndex        =   28
            Top             =   1620
            Width           =   615
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Eficiencia:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -74760
            TabIndex        =   26
            Top             =   3420
            Width           =   1020
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Reales:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -74760
            TabIndex        =   24
            Top             =   2940
            Width           =   750
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Estimadas:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -74760
            TabIndex        =   22
            Top             =   2460
            Width           =   1080
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Esperadas:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   -74760
            TabIndex        =   20
            Top             =   1980
            Width           =   1095
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "A:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -72600
            TabIndex        =   17
            Top             =   1140
            Width           =   240
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Semana:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -74760
            TabIndex        =   16
            Top             =   1140
            Width           =   825
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Operario:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   -74760
            TabIndex        =   15
            Top             =   780
            Width           =   960
         End
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   5040
         TabIndex        =   11
         Top             =   1800
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "MMMMMM/dddddd/yyyy"
         Format          =   16842753
         UpDown          =   -1  'True
         CurrentDate     =   39731
      End
      Begin Proyecto1.SComboBox cbooperario 
         Height          =   495
         Left            =   1680
         TabIndex        =   10
         Top             =   600
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   873
         AppearanceCombo =   16
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxListLength   =   -1
         NumberItemsToShow=   -1
         OfficeAppearance=   2
         ShadowColorText =   6582129
         Style           =   1
         Text            =   ""
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   1800
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "MMMMMM/dddddd/yyyy"
         Format          =   16842753
         UpDown          =   -1  'True
         CurrentDate     =   39730
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00EECB82&
         Caption         =   "Operario:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   1200
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00EECB82&
         Caption         =   "Semana:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   360
         TabIndex        =   7
         Top             =   1800
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00EECB82&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9600
         TabIndex        =   6
         Top             =   2640
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00EECB82&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   9480
         TabIndex        =   1
         Top             =   2160
         Visible         =   0   'False
         Width           =   75
      End
   End
   Begin HookMenu.XpMenu XpMenu1 
      Left            =   0
      Top             =   0
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   3
      CheckBorderColor=   0
      SelMenuBorder   =   7021576
      SelMenuBackColor=   14737632
      SelMenuForeColor=   12582912
      SelCheckBackColor=   13740436
      MenuBorderColor =   8388608
      SeparatorColor  =   16711680
      MenuBackColor   =   16711680
      MenuForeColor   =   0
      CheckBackColor  =   15326939
      CheckForeColor  =   10027263
      DisabledMenuBorderColor=   4210752
      DisabledMenuBackColor=   12632256
      DisabledMenuForeColor=   16711680
      MenuBarBackColor=   15649666
      MenuPopupBackColor=   15649666
      ShortCutNormalColor=   16711680
      ShortCutSelectColor=   12582912
      ArrowNormalColor=   15649666
      ArrowSelectColor=   16744576
      ShadowColor     =   0
      Bmp:1           =   "frmverreportes.frx":4759
      Key:1           =   "#mnaarchivosalir"
      Bmp:2           =   "frmverreportes.frx":4B81
      Key:2           =   "#mnuayudaayuda"
      Bmp:3           =   "frmverreportes.frx":4FA9
      Key:3           =   "#mnuarchivosalir"
      UseSystemFont   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "Eficiencia:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   47
      Top             =   1440
      Width           =   1020
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "Reales:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   46
      Top             =   960
      Width           =   750
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "Estimadas:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   45
      Top             =   480
      Width           =   1080
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Esperadas:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Width           =   1095
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuarchivosalir 
         Caption         =   "Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuayuda 
      Caption         =   "?"
      Begin VB.Menu mnuayudaayuda 
         Caption         =   "Ayuda"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmverreportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fecha, fecha2 As Date

Private Sub cmdgenerar_Click()
If sstreporte.Caption = "Semana" Then
Dim puntadas, tendido, a, n, uh, ud, horas As Long

If cbooperario.Text = "" Then
MsgBox ("Elija un operario"), vbExclamation, "Elejir"
Else
DTPicker1.Value = DateAdd("y", -6, DTPicker2.Value)
With mscreporte


'Produccion esperada
consultar "select * from reporte, operario where reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha between(#" & Label1.Caption & "#) and (#" & Label2.Caption & "#)"

If consulta.RecordCount = 0 Then
MsgBox ("No hay registro de ese operario en esa fecha"), vbExclamation, "Vacio"
Else

consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

.Column = 1
.Row = 1
.Data = ud
Else
.Column = 1
.Row = 1
.Data = 0
End If


DTPicker1.Value = DateAdd("y", 1, DTPicker1.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

.Column = 1
.Row = 2
.Data = ud
Else
.Column = 1
.Row = 2
.Data = 0
End If

DTPicker1.Value = DateAdd("y", 1, DTPicker1.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

.Column = 1
.Row = 3
.Data = ud
Else
.Column = 1
.Row = 3
.Data = 0
End If

DTPicker1.Value = DateAdd("y", 1, DTPicker1.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

.Column = 1
.Row = 4
.Data = ud
Else
.Column = 1
.Row = 4
.Data = 0
End If

DTPicker1.Value = DateAdd("y", 1, DTPicker1.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

.Column = 1
.Row = 5
.Data = ud
Else
.Column = 1
.Row = 5
.Data = 0
End If

DTPicker1.Value = DateAdd("y", 1, DTPicker1.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

.Column = 1
.Row = 6
.Data = ud
Else
.Column = 1
.Row = 6
.Data = 0
End If

DTPicker1.Value = DateAdd("y", 1, DTPicker1.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

.Column = 1
.Row = 7
.Data = ud
Else
.Column = 1
.Row = 7
.Data = 0
End If



DTPicker1.Value = DateAdd("y", -6, DTPicker2.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year

'Produccion estimada

consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and (causa.causa='normal' or causa.causa='problema')"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

consulta.MoveFirst

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

.Column = 2
.Row = 1
.Data = ud
Else
.Column = 2
.Row = 1
.Data = 0
End If


DTPicker1.Value = DateAdd("y", 1, DTPicker1.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

.Column = 2
.Row = 2
.Data = ud
Else
.Column = 2
.Row = 2
.Data = 0
End If

DTPicker1.Value = DateAdd("y", 1, DTPicker1.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

.Column = 2
.Row = 3
.Data = ud
Else
.Column = 2
.Row = 3
.Data = 0
End If

DTPicker1.Value = DateAdd("y", 1, DTPicker1.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

.Column = 2
.Row = 4
.Data = ud
Else
.Column = 2
.Row = 4
.Data = 0
End If

DTPicker1.Value = DateAdd("y", 1, DTPicker1.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

.Column = 2
.Row = 5
.Data = ud
Else
.Column = 2
.Row = 5
.Data = 0
End If

DTPicker1.Value = DateAdd("y", 1, DTPicker1.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

.Column = 2
.Row = 6
.Data = ud
Else
.Column = 2
.Row = 6
.Data = 0
End If

DTPicker1.Value = DateAdd("y", 1, DTPicker1.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

.Column = 2
.Row = 7
.Data = ud
Else
.Column = 2
.Row = 7
.Data = 0
End If



'Produccion real

DTPicker1.Value = DateAdd("y", -6, DTPicker2.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select fecha,cantidad from reporte, operario where reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha between(#" & Label1.Caption & "#) and (#" & Label2.Caption & "#)"
consulta.MoveFirst
.Column = 3
.Row = 1
Do Until consulta.EOF
If consulta.Fields("fecha") = DTPicker1.Value Then
.Data = consulta.Fields("cantidad")
consulta.MoveLast
 Else
.Data = 0
End If
consulta.MoveNext
Loop

consulta.MoveFirst

.Column = 3
.Row = 2
fecha2 = DateAdd("y", 1, DTPicker1.Value)

Do Until consulta.EOF
If consulta.Fields("fecha") = fecha2 Then
.Data = consulta.Fields("cantidad")
 consulta.MoveLast
 Else
.Data = 0
End If
consulta.MoveNext
Loop

consulta.MoveFirst

.Column = 3
.Row = 3
fecha2 = DateAdd("y", 2, DTPicker1.Value)

Do Until consulta.EOF
If consulta.Fields("fecha") = fecha2 Then

.Data = consulta.Fields("cantidad")
 consulta.MoveLast
 Else
.Data = 0
End If
consulta.MoveNext
Loop

consulta.MoveFirst

.Column = 3
.Row = 4
fecha2 = DateAdd("y", 3, DTPicker1.Value)

Do Until consulta.EOF
If consulta.Fields("fecha") = fecha2 Then
.Data = consulta.Fields("cantidad")
 consulta.MoveLast
 Else
.Data = 0
End If
consulta.MoveNext
Loop

consulta.MoveFirst

.Column = 3
.Row = 5
fecha2 = DateAdd("y", 4, DTPicker1.Value)

Do Until consulta.EOF
If consulta.Fields("fecha") = fecha2 Then
.Data = consulta.Fields("cantidad")
 consulta.MoveLast
 Else
.Data = 0
End If
consulta.MoveNext
Loop

consulta.MoveFirst


.Column = 3
.Row = 6
fecha2 = DateAdd("y", 5, DTPicker1.Value)

Do Until consulta.EOF
If consulta.Fields("fecha") = fecha2 Then
.Data = consulta.Fields("cantidad")
 consulta.MoveLast
 Else
.Data = 0
End If
consulta.MoveNext
Loop

consulta.MoveFirst

.Column = 3
.Row = 7
fecha2 = DateAdd("y", 6, DTPicker1.Value)

Do Until consulta.EOF
If consulta.Fields("fecha") = fecha2 Then
.Data = consulta.Fields("cantidad")
 consulta.MoveLast
 Else
.Data = 0
End If
consulta.MoveNext
Loop

End If
End With
End If
End If
If sstreporte.Caption = "General" Then
 cargardatos
End If
If sstreporte.Caption = "Rendimiento" Then
If cbooperario.Text = "" Then
 MsgBox ("elija un operario"), vbExclamation, "Elejir"
 Else
 cargarrendimiento
End If
End If
End Sub

Private Sub cmdsiguiente_Click()
fecha = DateAdd("y", 7, DTPicker1.Value)
DTPicker1.Value = fecha
DTPicker1.DayOfWeek = 2
DTPicker2.Value = DateAdd("y", 6, DTPicker1.Value)
DTPicker2.DayOfWeek = 1
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
Label2.Caption = DTPicker2.Month & "/" & DTPicker2.Day & "/" & DTPicker2.Year
End Sub

Private Sub cmdanterior_Click()
fecha = DateAdd("y", -7, DTPicker1.Value)
DTPicker1.Value = fecha
DTPicker1.DayOfWeek = 2
DTPicker2.Value = DateAdd("y", 6, DTPicker1.Value)
DTPicker2.DayOfWeek = 1
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
Label2.Caption = DTPicker2.Month & "/" & DTPicker2.Day & "/" & DTPicker2.Year
End Sub

Private Sub cmdactual_click()
fecha = Date
DTPicker1.Value = fecha
DTPicker1.DayOfWeek = 2
DTPicker2.Value = DateAdd("y", 6, DTPicker1.Value)
DTPicker2.DayOfWeek = 1
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
Label2.Caption = DTPicker2.Month & "/" & DTPicker2.Day & "/" & DTPicker2.Year
End Sub

Private Sub Form_Load()
Dim n, a As Long
DTPicker1.Value = Date
DTPicker1.DayOfWeek = 2
DTPicker2.Value = DateAdd("y", 6, DTPicker1.Value)
DTPicker2.DayOfWeek = 1
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
Label2.Caption = DTPicker2.Month & "/" & DTPicker2.Day & "/" & DTPicker2.Year
cargaroperario
With mscreporte
.Column = 1
.Row = 1
.Data = 0

.Column = 2
.Row = 1
.Data = 0

.Column = 3
.Row = 1
.Data = 0

.Column = 1
.Row = 2
.Data = 0

.Column = 2
.Row = 2
.Data = 0

.Column = 3
.Row = 2
.Data = 0

.Column = 1
.Row = 3
.Data = 0

.Column = 2
.Row = 3
.Data = 0

.Column = 3
.Row = 3
.Data = 0

.Column = 1
.Row = 4
.Data = 0

.Column = 2
.Row = 4
.Data = 0

.Column = 3
.Row = 4
.Data = 0

.Column = 1
.Row = 5
.Data = 0

.Column = 2
.Row = 5
.Data = 0

.Column = 3
.Row = 5
.Data = 0

.Column = 1
.Row = 6
.Data = 0

.Column = 2
.Row = 6
.Data = 0

.Column = 3
.Row = 6
.Data = 0

.Column = 1
.Row = 7
.Data = 0

.Column = 2
.Row = 7
.Data = 0

.Column = 3
.Row = 7
.Data = 0

End With

With mscreportegen
consultar "select count(codigo) as numero from operario"
.RowCount = 7
.ColumnCount = consulta.Fields("numero")
n = consulta.Fields("numero")

consultar "select nombre from operario"
If Not consulta.RecordCount = 0 Then
consulta.MoveLast


Do Until consulta.BOF
.Column = n
.ColumnLabel = consulta.Fields("nombre")
consulta.MovePrevious
n = n - 1
Loop
End If
End With
End Sub

Private Sub cargaroperario()
consultar "select nombre from operario"
If Not consulta.RecordCount = 0 Then
 Do Until consulta.EOF
  cbooperario.AddItem (consulta.Fields("nombre"))
  consulta.MoveNext
 Loop
End If
End Sub

Private Sub mnuarchivosalir_Click()
Unload Me
End Sub

Private Sub sstreporte_Click(PreviousTab As Integer)
If sstreporte.Caption = "General" Then
 cargardatos
End If
End Sub

Private Sub cargardatos()
Dim n As Long
consultar "select count(codigo) as numero from operario"
n = consulta.Fields("numero")


consultar2 "select * from operario"
If Not consulta2.RecordCount = 0 Then
consulta2.MoveLast

Do Until consulta2.BOF
With mscreportegen

DTPicker1.Value = DateAdd("y", -6, DTPicker2.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select fecha,cantidad from reporte, operario where reporte.operario=operario.codigo and operario.nombre='" & consulta2.Fields("nombre") & "'" & " and fecha between(#" & Label1.Caption & "#) and (#" & Label2.Caption & "#)"
If Not consulta.RecordCount = 0 Then
consulta.MoveFirst
.Column = n
.Row = 1
Do Until consulta.EOF
If consulta.Fields("fecha") = DTPicker1.Value Then
.Data = consulta.Fields("cantidad")
consulta.MoveLast
 Else
.Data = 0
End If
consulta.MoveNext
Loop

consulta.MoveFirst

.Column = n
.Row = 2
fecha2 = DateAdd("y", 1, DTPicker1.Value)

Do Until consulta.EOF
If consulta.Fields("fecha") = fecha2 Then
.Data = consulta.Fields("cantidad")
 consulta.MoveLast
 Else
.Data = 0
End If
consulta.MoveNext
Loop

consulta.MoveFirst

.Column = n
.Row = 3
fecha2 = DateAdd("y", 2, DTPicker1.Value)

Do Until consulta.EOF
If consulta.Fields("fecha") = fecha2 Then
.Data = consulta.Fields("cantidad")
 consulta.MoveLast
 Else
.Data = 0
End If
consulta.MoveNext
Loop

consulta.MoveFirst

.Column = n
.Row = 4
fecha2 = DateAdd("y", 3, DTPicker1.Value)

Do Until consulta.EOF
If consulta.Fields("fecha") = fecha2 Then
.Data = consulta.Fields("cantidad")
 consulta.MoveLast
 Else
.Data = 0
End If
consulta.MoveNext
Loop

consulta.MoveFirst

.Column = n
.Row = 5
fecha2 = DateAdd("y", 4, DTPicker1.Value)

Do Until consulta.EOF
If consulta.Fields("fecha") = fecha2 Then
.Data = consulta.Fields("cantidad")
 consulta.MoveLast
 Else
.Data = 0
End If
consulta.MoveNext
Loop

consulta.MoveFirst


.Column = n
.Row = 6
fecha2 = DateAdd("y", 5, DTPicker1.Value)

Do Until consulta.EOF
If consulta.Fields("fecha") = fecha2 Then
.Data = consulta.Fields("cantidad")
 consulta.MoveLast
 Else
.Data = 0
End If
consulta.MoveNext
Loop

consulta.MoveFirst

.Column = n
.Row = 7
fecha2 = DateAdd("y", 6, DTPicker1.Value)

Do Until consulta.EOF
If consulta.Fields("fecha") = fecha2 Then
.Data = consulta.Fields("cantidad")
 consulta.MoveLast
 Else
.Data = 0
End If
consulta.MoveNext
Loop

Else
.Column = n
.Row = 1
.Data = 0
.Column = n
.Row = 2
.Data = 0
.Column = n
.Row = 3
.Data = 0
.Column = n
.Row = 4
.Data = 0
.Column = n
.Row = 5
.Data = 0
.Column = n
.Row = 6
.Data = 0
.Column = n
.Row = 7
.Data = 0
End If
n = n - 1
consulta2.MovePrevious
End With
Loop
End If
End Sub
Private Sub cargarrendimiento()
lbloperario.Caption = cbooperario.Text

txtsemanade.Text = DTPicker1.Value
txtsemanaa.Text = DTPicker2.Value

DTPicker1.Value = DateAdd("y", -6, DTPicker2.Value)

'Produccion esperada
consultar "select * from reporte, operario where reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha between(#" & Label1.Caption & "#) and (#" & Label2.Caption & "#)"

If consulta.RecordCount = 0 Then
MsgBox ("No hay registro de ese operario en esa fecha"), vbExclamation, "Vacio"
txtesperadaslu.Text = 0
txtesperadasma.Text = 0
txtesperadasmi.Text = 0
txtesperadasju.Text = 0
txtesperadasvi.Text = 0
txtesperadassa.Text = 0
txtesperadasdo.Text = 0
txtestimadaslu.Text = 0
txtestimadasma.Text = 0
txtestimadasmi.Text = 0
txtestimadasju.Text = 0
txtestimadasvi.Text = 0
txtestimadassa.Text = 0
txtestimadasdo.Text = 0
txtrealeslu.Text = 0
txtrealesma.Text = 0
txtrealesmi.Text = 0
txtrealesju.Text = 0
txtrealesvi.Text = 0
txtrealessa.Text = 0
txtrealesdo.Text = 0
txteficiencialu.Text = 0
txteficienciama.Text = 0
txteficienciami.Text = 0
txteficienciaju.Text = 0
txteficienciavi.Text = 0
txteficienciasa.Text = 0
txteficienciado.Text = 0

Else

consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

txtesperadaslu.Text = ud
Else
txtesperadaslu = 0
End If


DTPicker1.Value = DateAdd("y", 1, DTPicker1.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

txtesperadasma.Text = ud
Else
txtesperadasma.Text = 0
End If

DTPicker1.Value = DateAdd("y", 1, DTPicker1.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

txtesperadasmi.Text = ud
Else
txtesperadasmi.Text = 0
End If

DTPicker1.Value = DateAdd("y", 1, DTPicker1.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

txtesperadasju.Text = ud
Else
txtesperadasju.Text = 0
End If

DTPicker1.Value = DateAdd("y", 1, DTPicker1.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

txtesperadasvi.Text = ud
Else
txtesperadasvi.Text = 0
End If

DTPicker1.Value = DateAdd("y", 1, DTPicker1.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

txtesperadassa.Text = ud
Else
txtesperadassa.Text = 0
End If

DTPicker1.Value = DateAdd("y", 1, DTPicker1.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and causa.causa='normal'"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

txtesperadasdo.Text = ud
Else
txtesperadasdo.Text = 0
End If


DTPicker1.Value = DateAdd("y", -6, DTPicker2.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year

'Produccion estimada

consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and (causa.causa='normal' or causa.causa='problema')"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

consulta.MoveFirst

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

txtestimadaslu.Text = ud
Else
txtestimadaslu.Text = 0
End If


DTPicker1.Value = DateAdd("y", 1, DTPicker1.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

txtestimadasma.Text = ud
Else
txtestimadasma.Text = 0
End If

DTPicker1.Value = DateAdd("y", 1, DTPicker1.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

txtestimadasmi.Text = ud
Else
txtestimadasmi.Text = 0
End If

DTPicker1.Value = DateAdd("y", 1, DTPicker1.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

txtestimadasju.Text = ud
Else
txtestimadasju.Text = 0
End If

DTPicker1.Value = DateAdd("y", 1, DTPicker1.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

txtestimadasvi.Text = ud
Else
txtestimadasvi.Text = 0
End If

DTPicker1.Value = DateAdd("y", 1, DTPicker1.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

txtestimadassa.Text = ud
Else
txtestimadassa.Text = 0
End If

DTPicker1.Value = DateAdd("y", 1, DTPicker1.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select count(reporte.reporte) as numero from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
If consulta.Fields("numero") = 0 Then
n = 0
Else
n = consulta.Fields("numero")
End If

If Not n = 0 Then
consultar "select * from reporte, produccion, reporte_detalle, tiempo_muerto, causa, operario where reporte.produccion=produccion.codigo and reporte.reporte=reporte_detalle.reporte and reporte_detalle.causa=tiempo_muerto.codigo and tiempo_muerto.causa=causa.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#" & " and  (causa.causa='normal' or causa.causa='problema')"
Else
consultar "select * from reporte, produccion, operario where reporte.produccion=produccion.codigo and reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha=#" & Label1.Caption & "#"
End If
If Not consulta.RecordCount = 0 Then
puntadas = consulta.Fields("cantidad_puntadas")
puntadas = puntadas / 300

If n = 0 Then
puntadas = puntadas
Else
For a = 1 To n
puntadas = puntadas + consulta.Fields("duracion")
consulta.MoveNext
Next
End If
tendido = puntadas
consulta.MoveFirst
uh = (60 / tendido) * consulta.Fields("cabezotes")
horas = Hour(consulta.Fields("hora_salida")) - Hour(consulta.Fields("hora_inicio"))
ud = uh * horas

txtestimadasdo.Text = ud
Else
txtestimadasdo.Text = 0
End If

'produccion real

DTPicker1.Value = DateAdd("y", -6, DTPicker2.Value)
Label1.Caption = DTPicker1.Month & "/" & DTPicker1.Day & "/" & DTPicker1.Year
consultar "select fecha,cantidad from reporte, operario where reporte.operario=operario.codigo and operario.nombre='" & cbooperario.Text & "'" & " and fecha between(#" & Label1.Caption & "#) and (#" & Label2.Caption & "#)"
consulta.MoveFirst
Do Until consulta.EOF
If consulta.Fields("fecha") = DTPicker1.Value Then
txtrealeslu.Text = consulta.Fields("cantidad")
consulta.MoveLast
 Else
txtrealeslu.Text = 0
End If
consulta.MoveNext
Loop

consulta.MoveFirst

fecha2 = DateAdd("y", 1, DTPicker1.Value)

Do Until consulta.EOF
If consulta.Fields("fecha") = fecha2 Then
txtrealesma.Text = consulta.Fields("cantidad")
 consulta.MoveLast
 Else
txtrealesma.Text = 0
End If
consulta.MoveNext
Loop

consulta.MoveFirst

fecha2 = DateAdd("y", 2, DTPicker1.Value)

Do Until consulta.EOF
If consulta.Fields("fecha") = fecha2 Then

txtrealesmi.Text = consulta.Fields("cantidad")
 consulta.MoveLast
 Else
txtrealesmi.Text = 0
End If
consulta.MoveNext
Loop

consulta.MoveFirst

fecha2 = DateAdd("y", 3, DTPicker1.Value)

Do Until consulta.EOF
If consulta.Fields("fecha") = fecha2 Then
txtrealesju.Text = consulta.Fields("cantidad")
 consulta.MoveLast
 Else
txtrealesju.Text = 0
End If
consulta.MoveNext
Loop

consulta.MoveFirst

fecha2 = DateAdd("y", 4, DTPicker1.Value)

Do Until consulta.EOF
If consulta.Fields("fecha") = fecha2 Then
txtrealesvi.Text = consulta.Fields("cantidad")
 consulta.MoveLast
 Else
txtrealesvi.Text = 0
End If
consulta.MoveNext
Loop

consulta.MoveFirst



fecha2 = DateAdd("y", 5, DTPicker1.Value)

Do Until consulta.EOF
If consulta.Fields("fecha") = fecha2 Then
txtrealessa.Text = consulta.Fields("cantidad")
 consulta.MoveLast
 Else
txtrealessa.Text = 0
End If
consulta.MoveNext
Loop

consulta.MoveFirst

fecha2 = DateAdd("y", 6, DTPicker1.Value)

Do Until consulta.EOF
If consulta.Fields("fecha") = fecha2 Then
txtrealesdo.Text = consulta.Fields("cantidad")
 consulta.MoveLast
 Else
txtrealesdo.Text = 0
End If
consulta.MoveNext
Loop

'eficiencia
If Not txtestimadaslu.Text = 0 Then
txteficiencialu.Text = (txtrealeslu.Text * 100) / txtestimadaslu.Text
Else
txteficiencialu.Text = 0
End If
If Not txtestimadasma.Text = 0 Then
txteficienciama.Text = (txtrealesma.Text * 100) / txtestimadasma.Text
Else
txteficienciama.Text = 0
End If
If Not txtestimadasmi.Text = 0 Then
txteficienciami.Text = (txtrealesmi.Text * 100) / txtestimadasmi.Text
Else
txteficienciami.Text = 0
End If
If Not txtestimadasju.Text = 0 Then
txteficienciaju.Text = (txtrealesju.Text * 100) / txtestimadasju.Text
Else
txteficienciaju.Text = 0
End If
If Not txtestimadasvi.Text = 0 Then
txteficienciavi.Text = (txtrealesvi.Text * 100) / txtestimadasvi.Text
Else
txteficienciavi.Text = 0
End If
If Not txtestimadassa.Text = 0 Then
txteficienciasa.Text = (txtrealessa.Text * 100) / txtestimadassa.Text
Else
txteficienciasa.Text = 0
End If
If Not txtestimadasdo.Text = 0 Then
txteficienciado.Text = (txtrealesdo.Text * 100) / txtestimadasdo.Text
Else
txteficienciado.Text = 0
End If

End If
End Sub
