VERSION 5.00
Object = "{4EC837EB-4201-4C6D-A7D7-0C99D69CD9A6}#1.0#0"; "GradientCommand.ocx"
Object = "{13592B48-653C-491D-ACB1-C3140AA12F33}#6.1#0"; "ubGrid.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmreportes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reportes"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   16515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   16515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   14208
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
      Begin GradientCommand.GGCommand cmdsalir 
         Height          =   615
         Left            =   11280
         TabIndex        =   10
         Top             =   4680
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
      Begin GradientCommand.GGCommand cmdguardart 
         Height          =   615
         Left            =   11280
         TabIndex        =   37
         Top             =   5520
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
      Begin GradientCommand.GGCommand cmdborrar 
         Height          =   615
         Left            =   11280
         TabIndex        =   38
         Top             =   6360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
         Caption         =   "Borrar"
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
      Begin XPFrame.FrameXp frxdetalle 
         Height          =   2055
         Left            =   6120
         TabIndex        =   13
         Top             =   4680
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   3625
         BackColor       =   15649666
         Caption         =   "Menu Detalle"
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
         Begin Proyecto1.SComboBox cbocausa 
            Height          =   450
            Left            =   1560
            TabIndex        =   43
            Top             =   840
            Width           =   2895
            _ExtentX        =   5106
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
         Begin Proyecto1.SComboBox cbotipo 
            Height          =   450
            Left            =   1560
            TabIndex        =   42
            Top             =   360
            Width           =   2895
            _ExtentX        =   5106
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
         Begin VB.TextBox txtduracion 
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
            Left            =   1560
            TabIndex        =   36
            Top             =   1320
            Width           =   1335
         End
         Begin VB.ComboBox cbocodigotipo 
            Height          =   315
            Left            =   960
            TabIndex        =   34
            Top             =   360
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.ComboBox cbocodigocausa 
            Height          =   315
            Left            =   960
            TabIndex        =   31
            Top             =   1680
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackColor       =   &H00EECB82&
            Caption         =   "Duración:"
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
            Left            =   120
            TabIndex        =   35
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackColor       =   &H00EECB82&
            Caption         =   "Motivo:"
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
            Left            =   120
            TabIndex        =   27
            Top             =   840
            Width           =   945
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00EECB82&
            Caption         =   "Causa:"
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
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   855
         End
      End
      Begin XPFrame.FrameXp frxdatos 
         Height          =   1815
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   16215
         _ExtentX        =   28601
         _ExtentY        =   3201
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
         Begin Proyecto1.SComboBox cbocabezotes 
            Height          =   450
            Left            =   8040
            TabIndex        =   45
            Top             =   360
            Width           =   1095
            _ExtentX        =   1931
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
         Begin Proyecto1.SComboBox cbomaquina 
            Height          =   450
            Left            =   4920
            TabIndex        =   41
            Top             =   1080
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
         Begin Proyecto1.SComboBox cboproduccion 
            Height          =   450
            Left            =   4920
            TabIndex        =   40
            Top             =   360
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
         Begin Proyecto1.SComboBox cbooperario 
            Height          =   450
            Left            =   1440
            TabIndex        =   39
            Top             =   1080
            Width           =   1815
            _ExtentX        =   3201
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
         Begin MSComCtl2.DTPicker dtpfecha 
            Height          =   405
            Left            =   11400
            TabIndex        =   32
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   714
            _Version        =   393216
            Format          =   50397185
            CurrentDate     =   39719
         End
         Begin VB.ComboBox cbocodigomaquina 
            Height          =   315
            Left            =   4320
            TabIndex        =   30
            Top             =   1440
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.ComboBox cbocodigoproduccion 
            Height          =   315
            Left            =   4200
            TabIndex        =   29
            Top             =   720
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.ComboBox cbocodigooperario 
            Height          =   315
            Left            =   720
            TabIndex        =   28
            Top             =   1440
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox txthorasal 
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
            Left            =   14760
            TabIndex        =   25
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txthoraent 
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
            Left            =   11400
            TabIndex        =   23
            Top             =   1080
            Width           =   1335
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
            Left            =   8040
            TabIndex        =   20
            Top             =   1080
            Width           =   1095
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
            TabIndex        =   14
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackColor       =   &H00EECB82&
            Caption         =   "Cabezotes:"
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
            Left            =   6600
            TabIndex        =   44
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H00EECB82&
            Caption         =   "Hora de Salida:"
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
            TabIndex        =   24
            Top             =   360
            Width           =   1905
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H00EECB82&
            Caption         =   "Hora de Entrada:"
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
            Left            =   9240
            TabIndex        =   22
            Top             =   1080
            Width           =   2145
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H00EECB82&
            Caption         =   "Fecha:"
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
            Left            =   9240
            TabIndex        =   21
            Top             =   360
            Width           =   825
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
            Left            =   6600
            TabIndex        =   19
            Top             =   1080
            Width           =   1200
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H00EECB82&
            Caption         =   "Maquina:"
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
            Left            =   3360
            TabIndex        =   18
            Top             =   1080
            Width           =   1170
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H00EECB82&
            Caption         =   "Produccion:"
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
            Left            =   3360
            TabIndex        =   17
            Top             =   360
            Width           =   1470
         End
         Begin VB.Label Label2 
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
            Left            =   120
            TabIndex        =   16
            Top             =   1080
            Width           =   1200
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
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   960
         End
      End
      Begin XPFrame.FrameXp frxmenudetalle 
         Height          =   1095
         Left            =   6120
         TabIndex        =   7
         Top             =   6840
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   1931
         BackColor       =   15649666
         Caption         =   "Menu Tiempo Muerto"
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
            Left            =   2040
            TabIndex        =   9
            Top             =   360
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
            Left            =   240
            TabIndex        =   8
            Top             =   360
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
      Begin XPFrame.FrameXp frxmenudatos 
         Height          =   4815
         Left            =   13320
         TabIndex        =   3
         Top             =   2280
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   8493
         BackColor       =   15649666
         Caption         =   "Menu Reportes"
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
         Begin GradientCommand.GGCommand cmdeliminar 
            Height          =   615
            Left            =   240
            TabIndex        =   11
            Top             =   4080
            Width           =   1575
            _ExtentX        =   2778
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
         Begin GradientCommand.GGCommand cmdmodificar 
            Height          =   615
            Left            =   240
            TabIndex        =   33
            Top             =   3240
            Width           =   1575
            _ExtentX        =   2778
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
         Begin GradientCommand.GGCommand cmdnuevo 
            Height          =   615
            Left            =   240
            TabIndex        =   4
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   1085
            Caption         =   "Nuevo"
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
         Begin GradientCommand.GGCommand cmdguardar 
            Height          =   615
            Left            =   240
            TabIndex        =   5
            Top             =   1320
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
         Begin GradientCommand.GGCommand cmdcancelar 
            Height          =   615
            Left            =   240
            TabIndex        =   6
            Top             =   2280
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
      End
      Begin ubGridControl.ubGrid fledetalle 
         Height          =   3255
         Left            =   240
         TabIndex        =   2
         Top             =   4680
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   5741
         Rows            =   1
         Cols            =   5
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
      Begin ubGridControl.ubGrid flereporte 
         Height          =   2175
         Left            =   240
         TabIndex        =   1
         Top             =   2280
         Width           =   12975
         _ExtentX        =   22886
         _ExtentY        =   3836
         Rows            =   1
         Cols            =   9
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
      Bmp:1           =   "frmreportes.frx":0000
      Key:1           =   "#mnaarchivosalir"
      Bmp:2           =   "frmreportes.frx":0428
      Key:2           =   "#mnuayudaayuda"
      Bmp:3           =   "frmreportes.frx":0850
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
   Begin VB.Menu mnuarchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuarchivosalir 
         Caption         =   "Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu Ayuda 
      Caption         =   "?"
      Begin VB.Menu mnuayudaayuda 
         Caption         =   "Ayuda"
      End
   End
End
Attribute VB_Name = "frmreportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, rt, cnt As Integer
Private Sub cargarflereporte()
consultar "select * from reporte, operario, produccion, maquina where reporte.operario=operario.codigo and reporte.produccion=produccion.codigo and reporte.maquina=maquina.codigo"


i = 1
With flereporte
.AutoSetup consulta.RecordCount, 9, True, True, "Codigo |Operario |Produccion |Maquina |Cabezotes |Cantidad |Fecha |Hora Entrada |Hora Salida"

If Not consulta.RecordCount = 0 Then
Do Until consulta.EOF
.Row = i
.Col = 0
.Text = i

.Col = 1
.Text = consulta.Fields("reporte")

.Col = 2
.Text = consulta.Fields("nombre")

.Col = 3
.Text = consulta.Fields("referencia")

.Col = 4
.Text = consulta.Fields("maquina.maquina")

.Col = 5
.Text = consulta.Fields("cabezotes")

.Col = 6
.Text = consulta.Fields("reporte.cantidad")

.Col = 7
.Text = consulta.Fields("fecha")

.Col = 8
.Text = consulta.Fields("hora_inicio")

.Col = 9
.Text = consulta.Fields("hora_salida")

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
End Sub

Private Sub cbocausa_SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
cbocodigocausa.ListIndex = cbocausa.ListIndex - 1
End Sub

Private Sub cbomaquina_SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
cbocodigomaquina.ListIndex = cbomaquina.ListIndex - 1
cargarcabezotes
End Sub

Private Sub cbooperario_SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
cbocodigooperario.ListIndex = cbooperario.ListIndex - 1
End Sub

Private Sub cboproduccion_SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
cbocodigoproduccion.ListIndex = cboproduccion.ListIndex - 1
End Sub

Private Sub cbotipo_SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
cbocodigotipo.ListIndex = cbotipo.ListIndex - 1
cargarcausa
End Sub

Private Sub cmdborrar_Click()
If rt = 1 Then
If MsgBox("Esta seguro de borrar la información?", vbQuestion + vbYesNo, "Borrar") = vbYes Then
 consultarotro "delete from  reporte where reporte=" & txtcodigo.Text
 consultarotro "update produccion set cantidad_faltantes= cantidad_faltantes+" & txtcantidad.Text & " where referencia='" & cboproduccion.Text & "'"
 consultar "select * from reporte_detalle where reporte=" & txtcodigo.Text
 If Not consulta.RecordCount = 0 Then
  consultarotro "delete from reporte_detalle where reporte=" & txtcodigo.Text
 End If
 txtcodigo.Text = ""
 cbooperario.Text = ""
 cbocodigooperario.Text = ""
 cboproduccion.Text = ""
 cbocodigoproduccion.Text = ""
 cbomaquina.Text = ""
 cbocabezotes.Text = ""
 cbocodigomaquina.Text = ""
 txtcantidad.Text = ""
 dtpfecha.Value = Date
 txthoraent.Text = ""
 txthorasal.Text = ""
 cargarflereporte
 habilitarini
End If
Else
MsgBox ("La información que usted trata de borrar procede de una modificación, esta información no puede ser borrada"), vbExclamation, "Error"
End If
End Sub

Private Sub cmdcancelar_Click()
cbooperario.Text = ""
cbocodigooperario.Text = ""
cboproduccion.Text = ""
cbocodigoproduccion.Text = ""
cbomaquina.Text = ""
cbocabezotes.Text = ""
cbocodigomaquina.Text = ""
txtcantidad.Text = ""
dtpfecha.Value = Date
txthoraent.Text = ""
txthorasal.Text = ""
habilitarini
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
st = 3
frmusuario.Show vbModal
If sh = 2 Then
 If MsgBox("Realmente desea eliminar esta produccion?", vbQuestion + vbYesNo, "Eliminar") = vbYes Then
  If Not txtcodigo.Text = "" Then
   consultar "select * from reporte where reporte=" & txtcodigo.Text
    If Not consulta.RecordCount = 0 Then
     consultarotro "delete from reporte where reporte=" & txtcodigo.Text
     cargarflereporte
      consultar "select * from reporte_detalle where reporte=" & txtcodigo.Text
       If Not consulta.RecordCount = 0 Then
       consultarotro "delete from reporte_detalle where reporte=" & txtcodigo.Text
       End If
       cmdcancelar_Click
    End If
  End If
 End If
End If
sh = 0
End Sub

Private Sub cmdguardar_Click()
If cbooperario.Text = "" Then
 MsgBox ("Escoja un operario"), vbExclamation, "Elejir"
 Else
 If cboproduccion.Text = "" Then
  MsgBox ("Escoja una producción"), vbExclamation, "Elejir"
  Else
  If cbomaquina.Text = "" Then
   MsgBox ("Escoja una maquina"), vbExclamation, "Elejir"
   Else
   If cbocabezotes.Text = "" Then
   MsgBox ("Escriba una cantidad de cabezotes"), vbExclamation, "Escribir"
   Else
   If txtcantidad.Text = "" Then
    MsgBox ("Escriba una cantidad"), vbExclamation, "Escribir"
    Else
    If txthoraent.Text = "" Then
     MsgBox ("Escriba una hora de entrada")
     Else
     If txthorasal.Text = "" Then
       MsgBox ("Escriba una hora de salida"), vbExclamation, "Escribir"
       Else
       If rt = 1 Then
       reporte.AddNew
       reporte.Fields("reporte") = txtcodigo.Text
       reporte.Fields("operario") = cbocodigooperario.Text
       reporte.Fields("maquina") = cbocodigomaquina.Text
       reporte.Fields("cabezotes") = cbocabezotes.Text
       reporte.Fields("produccion") = cbocodigoproduccion.Text
       reporte.Fields("cantidad") = txtcantidad.Text
       reporte.Fields("fecha") = dtpfecha.Value
       reporte.Fields("hora_inicio") = txthoraent.Text
       reporte.Fields("hora_salida") = txthorasal.Text
       reporte.Update
       cargarflereporte
       habilitargua
       consultarotro "update produccion set cantidad_faltantes= cantidad_faltantes-" & txtcantidad.Text & " where referencia='" & cboproduccion.Text & "'"
       Else
       consultarotro "update reporte set operario=" & cbocodigooperario.Text & ", maquina=" & cbocodigomaquina.Text & ", produccion=" & cbocodigoproduccion.Text & ", cantidad=" & txtcantidad.Text & ", fecha=#" & dtpfecha.Value & "#" & ", hora_inicio='" & txthoraent.Text & "'" & ", hora_salida='" & txthorasal.Text & "'" & " where reporte=" & txtcodigo.Text
       cargarflereporte
       habilitargua
       consultarotro "update produccion set cantidad_faltantes= cantidad_faltantes+" & cnt & " where referencia='" & cboproduccion.Text & "'"
       consultarotro "update produccion set cantidad_faltantes= cantidad_faltantes-" & txtcantidad.Text & " where referencia='" & cboproduccion.Text & "'"
     End If
    End If
   End If
  End If
 End If
 End If
End If
End If
End Sub

Private Sub cmdguardart_Click()
MsgBox ("Todo ha sido guardado con exito"), vbInformation, "Exito"
Unload Me
End Sub

Private Sub cmdmodificar_Click()
If Not txtcodigo.Text = "" Then
consultar "select * from produccion where codigo=" & cbocodigoproduccion.Text
If Not consulta.RecordCount = 0 Then
 rt = 2
 habilitarnue
 cnt = txtcantidad.Text
End If
Else
MsgBox ("Escoja una produccion para modificar"), vbExclamation, "Elejir"
End If
End Sub

Private Sub cmdnuevo_Click()
habilitarnue
consultar "select max(reporte) as reporte from reporte"

If Not consulta.RecordCount = 0 Then
If Not consulta.Fields("reporte") = 0 Then
txtcodigo.Text = consulta.Fields("reporte") + 1
Else
txtcodigo.Text = 1
End If
End If

cbooperario.Text = ""
cbocodigooperario.Text = ""
cboproduccion.Text = ""
cbocodigoproduccion.Text = ""
cbomaquina.Text = ""
cbocodigomaquina.Text = ""
cbocabezotes.Text = ""
txtcantidad.Text = ""
dtpfecha.Value = Date
txthoraent.Text = ""
txthorasal.Text = ""
rt = 1
fledetalle.Rows = 1
fledetalle.Row = 1
fledetalle.Col = 1
fledetalle.Text = ""
fledetalle.Col = 2
fledetalle.Text = ""
fledetalle.Col = 3
fledetalle.Text = ""
End Sub

Private Sub flereporte_DblClick()

If Not flereporte.Rows = 0 Then
txtcodigo.Text = flereporte.TextMatrix(flereporte.Row, 1)
cbooperario.Text = flereporte.TextMatrix(flereporte.Row, 2)
cboproduccion.Text = flereporte.TextMatrix(flereporte.Row, 3)
cbomaquina.Text = flereporte.TextMatrix(flereporte.Row, 4)
cbocabezotes.Text = flereporte.TextMatrix(flereporte.Row, 5)
txtcantidad.Text = flereporte.TextMatrix(flereporte.Row, 6)
dtpfecha.Value = flereporte.TextMatrix(flereporte.Row, 7)
txthoraent.Text = flereporte.TextMatrix(flereporte.Row, 8)
txthorasal.Text = flereporte.TextMatrix(flereporte.Row, 9)
consultar "select codigo from operario where nombre='" & cbooperario.Text & "'"
cbocodigooperario.Text = consulta.Fields("codigo")
consultar "select codigo from produccion where referencia='" & cboproduccion.Text & "'"
cbocodigoproduccion.Text = consulta.Fields("codigo")
consultar "select codigo from maquina where maquina='" & cbomaquina.Text & "'"
cbocodigomaquina.Text = consulta.Fields("codigo")

consultar "select * from reporte_detalle, causa, tiempo_muerto where tiempo_muerto.causa=causa.codigo and reporte_detalle.causa=tiempo_muerto.codigo and reporte=" & txtcodigo.Text
With fledetalle
If consulta.RecordCount = 0 Then

.Rows = 1
.Col = 1
.Text = ""
.Col = 2
.Text = ""
.Col = 3
.Text = ""
.Col = 4
.Text = ""
.Col = 5
.Text = ""
.Col = 6
.Text = ""
.Col = 7
.Text = ""
.Col = 8
.Text = ""
Else

.AutoSetup consulta.RecordCount, 3, True, True, "Causa |Motivo |Duración"
i = 1
Do Until consulta.EOF
 .Row = i
 .Col = 0
 .Text = i
 
 .Col = 1
 .Text = consulta.Fields("causa.causa")
 
 .Col = 2
 .Text = consulta.Fields("motivo")
 
 .Col = 3
 .Text = consulta.Fields("duracion")
 
 i = i + 1
 consulta.MoveNext
Loop
.ColAllowEdit(1) = False
.ColAllowEdit(2) = False
.ColAllowEdit(3) = False
End If
End With
End If
End Sub

Private Sub Form_Load()
abrirtablareporte
abrirtablareporte_detalle
cargarflereporte
habilitarini
cargaroperario
cargarproduccion
cargarmaquina
cargartipo
End Sub
Private Sub habilitarini()
frxmenudatos.Enabled = True
frxdatos.Enabled = False
frxdetalle.Enabled = False
frxmenudetalle.Enabled = False
fledetalle.Enabled = False
cmdguardar.Enabled = False
cmdcancelar.Enabled = False
cmdnuevo.Enabled = True
cmdmodificar.Enabled = True
cmdeliminar.Enabled = True
flereporte.Enabled = True
cmdguardart.Enabled = False
cmdborrar.Enabled = False
End Sub
Private Sub habilitarnue()
frxdatos.Enabled = True
cmdguardar.Enabled = True
cmdcancelar.Enabled = True
cmdnuevo.Enabled = False
cmdeliminar.Enabled = False
cmdmodificar.Enabled = False
flereporte.Enabled = False
End Sub
Private Sub habilitargua()
frxdatos.Enabled = False
flereporte.Enabled = False
frxdetalle.Enabled = True
fledetalle.Enabled = True
frxmenudetalle.Enabled = True
frxmenudatos.Enabled = False
cmdguardart.Enabled = True
cmdborrar.Enabled = True
End Sub
Private Sub cargaroperario()
consultar "select * from operario"

If Not consulta.RecordCount = 0 Then
Do Until consulta.EOF
cbooperario.AddItem (consulta.Fields("nombre"))
cbocodigooperario.AddItem (consulta.Fields("codigo"))
consulta.MoveNext
Loop
End If

End Sub
Private Sub cargarproduccion()
consultar "select * from produccion"

If Not consulta.RecordCount = 0 Then
Do Until consulta.EOF
cboproduccion.AddItem (consulta.Fields("referencia"))
cbocodigoproduccion.AddItem (consulta.Fields("codigo"))
consulta.MoveNext
Loop
End If

End Sub
Private Sub cargarmaquina()
consultar "select * from maquina"

If Not consulta.RecordCount = 0 Then
Do Until consulta.EOF
cbomaquina.AddItem (consulta.Fields("maquina"))
cbocodigomaquina.AddItem (consulta.Fields("codigo"))
consulta.MoveNext
Loop
End If

End Sub
Private Sub cargartipo()
consultar "select * from causa"

If Not consulta.RecordCount = 0 Then
Do Until consulta.EOF
cbotipo.AddItem (consulta.Fields("causa"))
cbocodigotipo.AddItem (consulta.Fields("codigo"))
consulta.MoveNext
Loop
End If

End Sub
Private Sub cargarcausa()
cbocausa.Clear
cbocodigocausa.Clear
consultar "select * from tiempo_muerto, causa where causa.codigo=tiempo_muerto.causa and tiempo_muerto.causa=" & cbocodigotipo.Text

If Not consulta.RecordCount = 0 Then
Do Until consulta.EOF
cbocausa.AddItem (consulta.Fields("motivo"))
cbocodigocausa.AddItem (consulta.Fields("tiempo_muerto.codigo"))
consulta.MoveNext
Loop
End If

End Sub
Private Sub cargarcabezotes()
Dim n As Integer
cbocabezotes.Clear
consultar "select cabezotes from tipo_maquina, maquina where maquina.tipo=tipo_maquina.codigo and maquina.maquina='" & cbomaquina.Text & "'"

If Not consulta.RecordCount = 0 Then
n = consulta.Fields("cabezotes")
Do Until n = 0
cbocabezotes.AddItem (n)
n = n - 1
Loop
End If

End Sub
Private Sub cargarfledetalle()
consultar "select * from reporte_detalle, causa, tiempo_muerto where tiempo_muerto.causa=causa.codigo and reporte_detalle.causa=tiempo_muerto.codigo and reporte=" & txtcodigo.Text
i = 1

With fledetalle
.AutoSetup consulta.RecordCount, 3, True, True, "Causa |Motivo |Duración"

If Not consulta.RecordCount = 0 Then
Do Until consulta.EOF
 .Row = i
 .Col = 0
 .Text = i
 
 .Col = 1
 .Text = consulta.Fields("causa.causa")
 
 .Col = 2
 .Text = consulta.Fields("motivo")
 
 .Col = 3
 .Text = consulta.Fields("duracion")
 
 i = i + 1
 consulta.MoveNext
Loop
.ColAllowEdit(1) = False
.ColAllowEdit(2) = False
.ColAllowEdit(3) = False
End If
End With
End Sub

Private Sub cmdagregar_Click()
If cbotipo.Text = "" Then
 MsgBox ("Escoja una causa"), vbExclamation, "Elejir"
 Else
 If cbocausa.Text = "" Then
  MsgBox ("Escoja un motivo"), vbExclamation, "Elejir"
  Else
  If txtduracion.Text = "" Then
   MsgBox ("Escriba una duracción"), vbExclamation, "Escribir"
   Else
   consultar "select * from reporte_detalle where reporte=" & txtcodigo.Text & " and causa=" & cbocodigocausa.Text
    If Not consulta.RecordCount = 0 Then
     MsgBox ("El motivo ya esta en la lista"), vbExclamation, "Error"
     Else
     reporte_detalle.AddNew
     reporte_detalle.Fields("reporte") = txtcodigo.Text
     reporte_detalle.Fields("causa") = cbocodigocausa.Text
     reporte_detalle.Fields("duracion") = txtduracion.Text
     reporte_detalle.Update
     cargarfledetalle
     cbocausa.Text = ""
     cbotipo.Text = ""
     cbocodigocausa.Text = ""
     cbocodigotipo.Text = ""
     txtduracion.Text = ""
    End If
  End If
 End If
End If
End Sub

Private Sub cmdquitar_Click()
If cbotipo.Text = "" Then
 MsgBox ("Escoja una causa"), vbExclamation, "Elejir"
 Else
 If cbocausa.Text = "" Then
  MsgBox ("Escoja un motivo"), vbExclamation, "Elejir"
  Else
  consultar "select * from reporte_detalle where reporte=" & txtcodigo.Text & " and causa=" & cbocodigocausa.Text
    If consulta.RecordCount = 0 Then
     MsgBox ("El motivo no esta en la lista"), vbExclamation, "Error"
     Else
     consultarotro "delete from reporte_detalle where reporte=" & txtcodigo.Text & " and causa=" & cbocodigocausa.Text
     cargarfledetalle
     cbocausa.Text = ""
     cbotipo.Text = ""
     cbocodigocausa.Text = ""
     cbocodigotipo.Text = ""
     txtduracion.Text = ""
    End If
 End If
End If
End Sub
Private Sub fledetalle_dblclick()
cbotipo.Text = fledetalle.TextMatrix(fledetalle.Row, 1)
cbocausa.Text = fledetalle.TextMatrix(fledetalle.Row, 2)
txtduracion.Text = fledetalle.TextMatrix(fledetalle.Row, 3)
consultar "select codigo from causa where causa='" & cbotipo.Text & "'"
cbocodigotipo.Text = consulta.Fields("codigo")
consultar "select codigo from tiempo_muerto where motivo='" & cbocausa.Text & "'"
cbocodigocausa.Text = consulta.Fields("codigo")
End Sub

Private Sub cmdsalir_Click()
If MsgBox("Si sale perdera todos los cambios no guardados, desea continuar?", vbQuestion + vbYesNo, "Cancelar") = vbYes Then
 Unload Me
End If
End Sub

Private Sub mnuarchivosalir_Click()
Unload Me
End Sub

Private Sub txtcantidad_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Then
 If keyascii = 13 Then
  txthoraent.SetFocus
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
