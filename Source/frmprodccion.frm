VERSION 5.00
Object = "{4EC837EB-4201-4C6D-A7D7-0C99D69CD9A6}#1.0#0"; "GradientCommand.ocx"
Object = "{13592B48-653C-491D-ACB1-C3140AA12F33}#6.1#0"; "ubGrid.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmproduccion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Producción"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   15600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   15600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin HookMenu.XpMenu XpMenu1 
      Left            =   1560
      Top             =   6360
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   2
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
      ShortCutSelectColor=   16711680
      ArrowNormalColor=   15649666
      ArrowSelectColor=   16744576
      ShadowColor     =   0
      Bmp:1           =   "frmprodccion.frx":0000
      Key:1           =   "#mnuarchivosalir"
      Bmp:2           =   "frmprodccion.frx":0428
      Key:2           =   "#mnuayudaayuda"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   9015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   15901
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
      Begin XPFrame.FrameXp FrameXp7 
         Height          =   1455
         Left            =   2520
         TabIndex        =   1
         Top             =   6120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   2566
         BackColor       =   15649666
         Caption         =   "Menu Tabla "
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
         Begin GradientCommand.GGCommand cmdver 
            Height          =   615
            Left            =   120
            TabIndex        =   2
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   1085
            Caption         =   "Ver Todos"
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
      End
      Begin XPFrame.FrameXp FrameXp6 
         Height          =   1455
         Left            =   4560
         TabIndex        =   3
         Top             =   6120
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   2566
         BackColor       =   15649666
         Caption         =   "Menu"
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
         Begin GradientCommand.GGCommand cmdsalir 
            Height          =   615
            Left            =   3720
            TabIndex        =   4
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   1085
            Caption         =   "Salir"
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
         Begin GradientCommand.GGCommand cmdcancelart 
            Height          =   615
            Left            =   1920
            TabIndex        =   5
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   1085
            Caption         =   "Cancelar"
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
         Begin GradientCommand.GGCommand cmdguardart 
            Height          =   615
            Left            =   120
            TabIndex        =   6
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   1085
            Caption         =   "Guardar"
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
      End
      Begin XPFrame.FrameXp FrameXp5 
         Height          =   1335
         Left            =   10320
         TabIndex        =   7
         Top             =   3600
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   2355
         BackColor       =   15649666
         Caption         =   "Menu Hilos"
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
         Begin GradientCommand.GGCommand cmdquitar 
            Height          =   615
            Left            =   1800
            TabIndex        =   8
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   1085
            Caption         =   "Quitar"
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
         Begin GradientCommand.GGCommand cmdagregar 
            Height          =   615
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   1085
            Caption         =   "Agregar"
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
      End
      Begin ubGridControl.ubGrid flehilo 
         Height          =   2415
         Left            =   10320
         TabIndex        =   10
         Top             =   5160
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   4260
         Rows            =   1
         Cols            =   2
         Redraw          =   -1  'True
         ShowGrid        =   -1  'True
         GridSolid       =   -1  'True
         GridLineColor   =   12632256
         BackColorFixed  =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontEdit {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ubGridControl.ubGrid fleproduccion 
         Height          =   2415
         Left            =   240
         TabIndex        =   11
         Top             =   3600
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   4260
         Rows            =   1
         Cols            =   8
         Redraw          =   -1  'True
         ShowGrid        =   -1  'True
         GridSolid       =   -1  'True
         GridLineColor   =   12632256
         BackColorFixed  =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontEdit {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XPFrame.FrameXp FrameXp4 
         Height          =   1335
         Left            =   9360
         TabIndex        =   12
         Top             =   2160
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2355
         BackColor       =   15649666
         Caption         =   "Hilos"
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
         Begin VB.ComboBox cbocodigohilo 
            Height          =   315
            Left            =   5280
            TabIndex        =   13
            Top             =   240
            Visible         =   0   'False
            Width           =   495
         End
         Begin Proyecto1.SComboBox cbocantidad 
            Height          =   450
            Left            =   4920
            TabIndex        =   14
            Top             =   600
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   794
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
         Begin Proyecto1.SComboBox cbohilo 
            Height          =   450
            Left            =   960
            TabIndex        =   15
            Top             =   600
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   794
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
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00EECB82&
            Caption         =   "Hilo:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   240
            TabIndex        =   17
            Top             =   600
            Width           =   585
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00EECB82&
            Caption         =   "Cantidad:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3480
            TabIndex        =   16
            Top             =   600
            Width           =   1200
         End
      End
      Begin XPFrame.FrameXp FrameXp3 
         Height          =   1335
         Left            =   240
         TabIndex        =   18
         Top             =   2160
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   2355
         BackColor       =   15649666
         Caption         =   "Menu Datos"
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
         Begin GradientCommand.GGCommand cmdnuevo 
            Height          =   615
            Left            =   120
            TabIndex        =   19
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   1085
            Caption         =   "Nuevo"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientToColor =   15649666
         End
         Begin GradientCommand.GGCommand cmdguardar 
            Height          =   615
            Left            =   3480
            TabIndex        =   20
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   1085
            Caption         =   "Guardar"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientToColor =   15649666
         End
         Begin GradientCommand.GGCommand cmdcancelar 
            Height          =   615
            Left            =   5160
            TabIndex        =   21
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   1085
            Caption         =   "Cancelar"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientToColor =   15649666
         End
         Begin GradientCommand.GGCommand cmdmodificar 
            Height          =   615
            Left            =   1800
            TabIndex        =   40
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   1085
            Caption         =   "Modificar"
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
         Begin GradientCommand.GGCommand cmdeliminar 
            Height          =   615
            Left            =   6840
            TabIndex        =   41
            Top             =   480
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   1085
            Caption         =   "Eliminar"
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
      End
      Begin XPFrame.FrameXp FrameXp2 
         Height          =   1695
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   2990
         BackColor       =   15649666
         Caption         =   "Datos"
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
         Begin VB.TextBox txtref 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4800
            TabIndex        =   42
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtcodigo 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   1320
            TabIndex        =   31
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox txtnombre 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   4800
            TabIndex        =   29
            Top             =   360
            Width           =   1815
         End
         Begin VB.TextBox txtcantidad 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   8280
            TabIndex        =   28
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtpuntadas 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   8280
            TabIndex        =   27
            Top             =   960
            Width           =   1335
         End
         Begin VB.TextBox txtfaltantes 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   11160
            TabIndex        =   26
            Top             =   960
            Width           =   1335
         End
         Begin VB.ComboBox cbocodigocliente 
            Height          =   315
            Left            =   2520
            TabIndex        =   23
            Top             =   480
            Visible         =   0   'False
            Width           =   615
         End
         Begin MSComCtl2.DTPicker dtpfin 
            Height          =   375
            Left            =   13440
            TabIndex        =   24
            Top             =   960
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "MM/dd/yyyy"
            Format          =   16777219
            CurrentDate     =   39711
            MinDate         =   -657434
         End
         Begin MSComCtl2.DTPicker dtpingreso 
            Height          =   405
            Left            =   11160
            TabIndex        =   25
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   714
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "MM/dd/yyyy"
            Format          =   16777219
            CurrentDate     =   39711
         End
         Begin Proyecto1.SComboBox cbocliente 
            Height          =   450
            Left            =   1320
            TabIndex        =   30
            Top             =   960
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   794
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
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00EECB82&
            Caption         =   "Referencia:"
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
            Left            =   3240
            TabIndex        =   43
            Top             =   960
            Width           =   1410
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00EECB82&
            Caption         =   "Codigo:"
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
            Left            =   240
            TabIndex        =   39
            Top             =   360
            Width           =   960
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00EECB82&
            Caption         =   "Cliente:"
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
            Left            =   240
            TabIndex        =   38
            Top             =   960
            Width           =   945
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00EECB82&
            Caption         =   "Descripción:"
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
            Left            =   3240
            TabIndex        =   37
            Top             =   360
            Width           =   1530
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00EECB82&
            Caption         =   "# Puntadas:"
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
            Left            =   6720
            TabIndex        =   36
            Top             =   960
            Width           =   1440
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H00EECB82&
            Caption         =   "Cantidad:"
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
            Left            =   6720
            TabIndex        =   35
            Top             =   360
            Width           =   1200
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00EECB82&
            Caption         =   "Faltante:"
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
            Left            =   9960
            TabIndex        =   34
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00EECB82&
            Caption         =   "Ingreso:"
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
            Left            =   9960
            TabIndex        =   33
            Top             =   360
            Width           =   1005
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00EECB82&
            Caption         =   "Fin:"
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
            Left            =   12840
            TabIndex        =   32
            Top             =   960
            Width           =   480
         End
      End
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
Attribute VB_Name = "frmproduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rt As Integer
Private Sub cbocliente_SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
cbocodigocliente.ListIndex = cbocliente.ListIndex - 1
txtnombre.SetFocus
End Sub

Private Sub cbohilo_SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
cbocodigohilo.ListIndex = cbohilo.ListIndex - 1
cargarcantidad
cbocantidad.SetFocus
End Sub
Private Sub cmdagregar_Click()
If cbohilo.Text = "" Then
 MsgBox ("Elija un hilo"), vbExclamation, "Elejir"
 Else
  If cbocantidad.Text = "" Then
   MsgBox ("Elija una cantidad"), vbExclamation, "Elejir"
   Else
    consultar "select color from hilo ,produccion_hilo where produccion_hilo.hilo=hilo.codigo and color='" & cbohilo.Text & "'" & " and produccion_hilo.codigo=" & txtcodigo.Text
    If consulta.RecordCount <> 0 Then
     MsgBox ("El hilo ya esta en la lista"), vbExclamation, "Existe"
     Else
      produccion_hilo.AddNew
      produccion_hilo.Fields("codigo") = txtcodigo.Text
      produccion_hilo.Fields("hilo") = cbocodigohilo.Text
      produccion_hilo.Fields("cantidad") = cbocantidad.Text
      produccion_hilo.Update
      cargarflehilo
      consultarotro "update hilo set cantidad=cantidad-" & cbocantidad.Text & " where color='" & cbohilo.Text & "'"
      cargarcantidad
      cargarhilo
    End If
  End If
End If
End Sub
Private Sub cargarflehilo()
Dim i As Integer
i = 1
consultar "select * from produccion_hilo, hilo where hilo.codigo = produccion_hilo.hilo and produccion_hilo.codigo=" & txtcodigo.Text
If consulta.RecordCount <> 0 Then
With flehilo
.AutoSetup consulta.RecordCount + 1, 2, True, True, "Color |Cantidad"
Do Until consulta.EOF
.Row = i
.Col = 0
.Text = i

.Col = 1
.Text = consulta.Fields("color")

.Col = 2
.Text = consulta.Fields("produccion_hilo.cantidad")

consulta.MoveNext
i = i + 1
Loop
.ColAllowEdit(1) = False
.ColAllowEdit(2) = False
End With
End If
autoajustarh
End Sub
Private Sub cmdcancelar_Click()
habilitarini
txtcodigo.Text = ""
txtnombre.Text = ""
cbocliente.Text = ""
cbocodigocliente.Text = ""
txtcantidad.Text = ""
txtpuntadas.Text = ""
txtfaltantes.Text = ""
dtpingreso.Value = CDate(Date)
dtpfin.Value = CDate(Date)
End Sub
Private Sub cmdcancelart_Click()
If Not txtcodigo.Text = "" Then
If MsgBox("Desea cancelar esta operación?", vbQuestion + vbYesNo, "Cancelar") = vbYes Then
consultarotro "delete from produccion where codigo=" & txtcodigo.Text
consultarotro "delete from produccion_hilo where codigo=" & txtcodigo.Text
dtpingreso.MinDate = Date
dtpfin.MinDate = dtpingreso.Value
txtnombre.Text = ""
txtref.Text = ""
cbocliente.Text = ""
cbocodigocliente.Text = ""
txtcantidad.Text = ""
txtpuntadas.Text = ""
txtfaltantes.Text = ""
dtpingreso.MinDate = Date
dtpfin.MinDate = Date
dtpingreso.Value = CDate(Date)
dtpfin.Value = CDate(Date)
txtcodigo.Text = ""
habilitarini
End If
End If
End Sub

Private Sub cmdeliminar_Click()
sh = 0
If consulta.State = 1 Then
 consulta.Close
End If
sentencia = "select * from usuario where tipo='Administrador'"
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic
frmusuario.cbonombre.Clear
Do Until consulta.EOF
 frmusuario.cbonombre.AddItem (consulta.Fields("nombre"))
 consulta.MoveNext
Loop
st = 2
frmusuario.Show vbModal
If sh = 1 Then
 If MsgBox("Realmente desea eliminar esta produccion?", vbQuestion + vbYesNo, "Eliminar") = vbYes Then
  If Not txtcodigo.Text = "" Then
   consultar "select * from produccion where codigo=" & txtcodigo.Text
    If Not consulta.RecordCount = 0 Then
     consultarotro "delete from produccion where codigo=" & txtcodigo.Text
     cargarfleproduccion
      consultar "select * from produccion_hilo where codigo=" & txtcodigo.Text
       If Not consulta.RecordCount = 0 Then
       consultarotro "delete from produccion_hilo where codigo=" & txtcodigo.Text
       End If
    End If
  End If
 End If
End If
cargarfleproduccion
sh = 0
End Sub

Private Sub cmdguardart_Click()
If txtcodigo.Text <> "" Then
consultar "select * from produccion where codigo=" & txtcodigo.Text
If consulta.RecordCount <> 0 Then
consultar "select * from produccion_hilo where codigo=" & txtcodigo.Text
If consulta.RecordCount = 0 Then
 MsgBox ("Aregue por lo menos un hilo")
Else
MsgBox ("Todo ha sido guardado exitosamente"), vbInformation, "Exito"
Unload Me
End If
End If
End If
End Sub

Private Sub cmdmodificar_Click()
If Not txtcodigo.Text = "" Then
consultar "select * from produccion where codigo=" & txtcodigo.Text
If Not consulta.RecordCount = 0 Then
 rt = 2
 habilitarnue
End If
Else
MsgBox ("Escoja una produccion para modificar"), vbExclamation, "Elejir"
End If
End Sub

Private Sub cmdnuevo_Click()
habilitarnue
dtpingreso.MinDate = Date
dtpfin.MinDate = dtpingreso.Value
txtnombre.Text = ""
cbocliente.Text = ""
cbocodigocliente.Text = ""
txtcantidad.Text = ""
txtpuntadas.Text = ""
txtfaltantes.Text = ""
txtref.Text = ""
dtpingreso.MinDate = Date
dtpfin.MinDate = Date
dtpingreso.Value = CDate(Date)
dtpfin.Value = CDate(Date)
With flehilo
.Rows = 1
i = 1
.Row = i
.Col = 0
.Text = i

.Col = 1
.Text = ""

.Col = 2
.Text = ""

consulta.MoveNext
i = i + 1

End With
rt = 1
End Sub
Private Sub cmdquitar_Click()
If cbohilo.Text = "" Then
MsgBox ("De doble click sobre algun hilo"), vbCritical, "Error"
Else
 If cbocantidad.Text = "" Then
  MsgBox ("De doble click sobre algun hilo"), vbCritical, "Error"
  Else
If MsgBox("Esta seguro de quitar este hilo?", vbQuestion + vbYesNo, "Quitar") = vbYes Then
 consultarotro "Delete from produccion_hilo where hilo=" & cbocodigohilo.Text & " and produccion_hilo.codigo=" & txtcodigo.Text
 consultarotro "update hilo set cantidad=cantidad+" & cbocantidad.Text & " where color='" & cbohilo.Text & "'"
 cargarflehilo
 End If
 cbohilo.Text = ""
End If
End If
End Sub
Private Sub cmdsalir_Click()
If MsgBox("Al salir se perderan todos los cambios que no hayan sido guardados, desea continuar?", vbQuestion + vbYesNo, "Salir") = vbYes Then
 If txtcodigo.Text <> "" Then
  consultar "select * from produccion where codigo=" & txtcodigo.Text
  If consulta.RecordCount <> 0 Then
  consultarotro "delete from produccion where codigo=" & txtcodigo.Text
  consultarotro "delete from produccion_hilo where codigo=" & txtcodigo.Text
 End If
 End If
 Unload Me
End If
End Sub

Private Sub cmdver_Click()
cargarfleproduccion
End Sub
Private Sub flehilo_DblClick()
cbohilo.Text = flehilo.TextMatrix(flehilo.Row, 1)
cbocantidad.Text = flehilo.TextMatrix(flehilo.Row, 2)
consultar "select hilo from produccion_hilo, hilo where hilo.codigo=produccion_hilo.hilo and color='" & cbohilo.Text & "'"
cbocodigohilo.Text = consulta.Fields("hilo")
End Sub

Private Sub fleproduccion_dblClick()
Dim i As Integer
If Not fleproduccion.Rows = 0 Then
txtcodigo.Text = fleproduccion.TextMatrix(fleproduccion.Row, 1)
txtnombre.Text = fleproduccion.TextMatrix(fleproduccion.Row, 3)
cbocliente.Text = fleproduccion.TextMatrix(fleproduccion.Row, 2)
txtref.Text = fleproduccion.TextMatrix(fleproduccion.Row, 4)
consultar "select cliente from produccion where codigo=" & txtcodigo.Text
cbocodigocliente.Text = consulta.Fields("cliente")
cerrarconsulta
txtpuntadas.Text = fleproduccion.TextMatrix(fleproduccion.Row, 7)
txtcantidad.Text = fleproduccion.TextMatrix(fleproduccion.Row, 5)
txtfaltantes.Text = fleproduccion.TextMatrix(fleproduccion.Row, 6)
dtpingreso.Value = CDate(fleproduccion.TextMatrix(fleproduccion.Row, 8))
dtpfin.Value = CDate(fleproduccion.TextMatrix(fleproduccion.Row, 9))
consultar "select * from produccion_hilo,hilo where hilo.codigo=produccion_hilo.hilo and produccion_hilo.codigo=" & txtcodigo.Text
With flehilo
.Rows = consulta.RecordCount + 1
i = 1
 Do Until consulta.EOF
.Row = i
.Col = 0
.Text = i

.Col = 1
.Text = consulta.Fields("color")

.Col = 2
.Text = consulta.Fields("produccion_hilo.cantidad")

consulta.MoveNext
i = i + 1
Loop
End With
End If
End Sub
Private Sub Form_Load()
abrirtablaproduccion
abrirtablaproduccion_hilo
abrirtablahilo
cargaricono Me
habilitarini
cargarfleproduccion
cargarcliente
rt = 0
End Sub

Private Sub mnuarchivosalir_Click()
Unload Me
End Sub
Private Sub habilitarini()
FrameXp2.Enabled = False
FrameXp4.Enabled = False
FrameXp5.Enabled = False
cmdnuevo.Enabled = True
cmdguardar.Enabled = False
cmdcancelar.Enabled = False
fleproduccion.Enabled = True
dtpingreso.MinDate = CDate("01/01/2001")
dtpfin.MinDate = CDate("01/01/2001")
FrameXp7.Enabled = True
cmdmodificar.Enabled = True
cmdeliminar.Enabled = True
End Sub
Private Sub habilitarnue()
If rt <> 2 Then
cargarcodigo "produccion", Me
End If
FrameXp2.Enabled = True
cmdguardar.Enabled = True
cmdcancelar.Enabled = True
cmdnuevo.Enabled = False
fleproduccion.Enabled = False
FrameXp7.Enabled = False
cmdmodificar.Enabled = False
cmdeliminar.Enabled = False
End Sub
Private Sub habilitargua()
FrameXp2.Enabled = False
cmdguardar.Enabled = False
cmdcancelar.Enabled = False
cmdnuevo.Enabled = False
FrameXp4.Enabled = True
FrameXp5.Enabled = True
End Sub
Private Sub cargarfleproduccion()
consultar "select * from produccion, cliente where produccion.cliente = cliente.codigo"
With fleproduccion
 .AutoSetup consulta.RecordCount, 9, True, True, "Codigo |Cliente |Descripción |Referencia |Cantidad |Faltantes |# Puntadas |Fecha Ingreso |Fecha Fin|"
 
If Not consulta.RecordCount = 0 Then
 i = 1
 consulta.MoveFirst
 Do Until consulta.EOF
 .Row = i
 .Col = 0
 .Text = i
 
 .Col = 1
 .Text = consulta.Fields("produccion.codigo")
 
 .Col = 2
 .Text = consulta.Fields("cliente.cliente")
 
 .Col = 3
 .Text = consulta.Fields("descripcion")
 
 .Col = 4
 .Text = consulta.Fields("referencia")
 
 .Col = 5
 .Text = consulta.Fields("cantidad")
 
 .Col = 6
 .Text = consulta.Fields("cantidad_faltantes")
 
 .Col = 7
 .Text = consulta.Fields("cantidad_puntadas")
 
 .Col = 8
 .Text = consulta.Fields("fecha_de_ingreso")
 
 .Col = 9
 .Text = consulta.Fields("fecha_salida")
 
 i = i + 1
 consulta.MoveNext
Loop
End If

.ColAllowEdit(1) = False
.ColAllowEdit(2) = False
.ColAllowEdit(3) = False
.ColAllowEdit(4) = False
.ColAllowEdit(5) = False
.ColAllowEdit(6) = False
.ColAllowEdit(7) = False
.ColAllowEdit(8) = False
.ColAllowEdit(9) = False
End With
autoajustar
End Sub
Private Sub cargarcliente()
consultar "select * from cliente"

If Not consulta.RecordCount = 0 Then
Do Until consulta.EOF
 cbocliente.AddItem (consulta.Fields("cliente"))
 cbocodigocliente.AddItem (consulta.Fields("codigo"))
 consulta.MoveNext
Loop
End If
End Sub
Private Sub cmdguardar_Click()
If rt = 1 Then
If txtnombre.Text = "" Then
 MsgBox ("Escriba un nombre para la produccion"), vbExclamation, "Escribir"
 Else
 If txtref.Text = "" Then
  MsgBox ("Escriba una referencia"), vbExclamation, "Escribir"
  Else
 If cbocliente.Text = "" Then
  MsgBox ("Elija un cliente"), vbExclamation, "Elejir"
  Else
  If txtpuntadas.Text = "" Then
   MsgBox ("Escriba una cantidad"), vbExclamation, "Escribir"
   Else
   If txtcantidad.Text = "" Then
    MsgBox ("Escriba una cantidad"), vbExclamation, "Escribir"
    Else
    produccion.AddNew
    produccion.Fields("codigo") = txtcodigo.Text
    produccion.Fields("cliente") = cbocodigocliente.Text
    produccion.Fields("descripcion") = txtnombre.Text
    produccion.Fields("referencia") = txtref.Text
    produccion.Fields("cantidad_puntadas") = txtpuntadas.Text
    produccion.Fields("cantidad") = txtcantidad.Text
    produccion.Fields("cantidad_faltantes") = txtcantidad.Text
    produccion.Fields("fecha_de_ingreso") = dtpingreso.Value
    produccion.Fields("fecha_salida") = dtpfin.Value
    habilitargua
    produccion.Update
    With fleproduccion
    .Rows = 1
    
    .Row = 1
    .Col = 1
    .Text = txtcodigo.Text
    
    .Col = 2
    .Text = cbocliente.Text
    
    .Col = 3
    .Text = txtnombre.Text
    
    .Col = 4
    .Text = txtref.Text
    
    .Col = 5
    .Text = txtcantidad.Text
    
    .Col = 6
    .Text = txtfaltantes.Text
    
    .Col = 7
    .Text = txtpuntadas.Text
    
    .Col = 8
    .Text = dtpingreso.Value
    
    .Col = 9
    .Text = dtpfin.Value
    cargarhilo
End With
End If
End If
End If
End If
End If
Else
 If rt = 2 Then
  consultarotro "update produccion set cliente=" & cbocodigocliente.Text & ", descripcion='" & txtnombre.Text & "'" & ", referencia='" & txtref.Text & "'" & ", cantidad=" & txtcantidad.Text & ", cantidad_puntadas=" & txtpuntadas.Text & ", cantidad_faltantes=cantidad_faltantes+(cantidad_faltantes-" & txtfaltantes.Text & ")" & ", fecha_de_ingreso=#" & dtpingreso.Value & "#" & " where codigo=" & txtcodigo.Text
  habilitargua
  produccion.Update
  With fleproduccion
    .Rows = 1
    .Row = 1
    .Col = 1
    .Text = txtcodigo.Text
    
    .Col = 2
    .Text = cbocliente.Text
    
    .Col = 3
    .Text = txtnombre.Text
    
    .Col = 4
    .Text = txtref.Text
    
    .Col = 5
    .Text = txtcantidad.Text
    
    .Col = 6
    .Text = txtfaltantes.Text
    
    .Col = 7
    .Text = txtpuntadas.Text
    
    .Col = 8
    .Text = dtpingreso.Value
    
    .Col = 9
    .Text = dtpfin.Value
    cargarhilo
End With
End If
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
Private Sub cargarhilo()
cbohilo.Clear
consultar "Select * from hilo"

If Not consulta.RecordCount = 0 Then
Do Until consulta.EOF
 cbohilo.AddItem (consulta.Fields("color"))
 cbocodigohilo.AddItem (consulta.Fields("codigo"))
 consulta.MoveNext
Loop
End If

End Sub
Private Sub cargarcantidad()
Dim c As Long
consultar "Select cantidad from hilo where color='" & cbohilo.Text & "'"

If Not consulta.RecordCount = 0 Then
c = consulta.Fields("cantidad")
cbocantidad.Clear
Do Until c = 0
 cbocantidad.AddItem (c)
 c = c - 1
Loop
End If

End Sub
Private Sub txtnombre_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Or keyascii = 241 Or keyascii = 209 Then
 If keyascii = 13 Then
 txtpuntadas.SetFocus
keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo letras y numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub txtref_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Or keyascii = 241 Or keyascii = 209 Then
 If keyascii = 13 Then
 txtpuntadas.SetFocus
keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo letras y numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub txtpuntadas_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Then
 If keyascii = 13 Then
  txtcantidad.SetFocus
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub txtcantidad_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Then
 If keyascii = 13 Then
  dtpingreso.SetFocus
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub dtpingreso_keypress(keyascii As Integer)
 If keyascii = 13 Then
  cmdguardar_Click
  keyascii = 0
 End If
End Sub
Private Sub autoajustar()
With fleproduccion
.ColWidth(1) = 50
.ColWidth(2) = 100
.ColWidth(3) = 200
.ColWidth(4) = 100
.ColWidth(5) = 70
.ColWidth(6) = 70
.ColWidth(7) = 100
.ColWidth(8) = 70
.ColWidth(9) = 70
End With
End Sub
Private Sub autoajustarh()
With flehilo
.ColWidth(1) = 100
.ColWidth(2) = 50
End With
End Sub
