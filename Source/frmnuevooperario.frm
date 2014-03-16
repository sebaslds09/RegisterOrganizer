VERSION 5.00
Object = "{4EC837EB-4201-4C6D-A7D7-0C99D69CD9A6}#1.0#0"; "GradientCommand.ocx"
Object = "{13592B48-653C-491D-ACB1-C3140AA12F33}#6.1#0"; "ubGrid.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Begin VB.Form frmnuevooperario 
   BackColor       =   &H00EECB82&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Operario"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   9551
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
      Begin TabDlg.SSTab sstoperario 
         Height          =   4575
         Left            =   0
         TabIndex        =   1
         Top             =   480
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   8070
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         Tab             =   3
         TabsPerRow      =   4
         TabHeight       =   520
         BackColor       =   15649666
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Nuevo"
         TabPicture(0)   =   "frmnuevooperario.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "cmdcancelar"
         Tab(0).Control(1)=   "cmdguardar"
         Tab(0).Control(2)=   "txtapellidos"
         Tab(0).Control(3)=   "RichTextBox1"
         Tab(0).Control(4)=   "txtdireccion"
         Tab(0).Control(5)=   "txttelefono"
         Tab(0).Control(6)=   "txtcedula"
         Tab(0).Control(7)=   "txtnombre"
         Tab(0).Control(8)=   "txtcodigo"
         Tab(0).Control(9)=   "Label6"
         Tab(0).Control(10)=   "Label5"
         Tab(0).Control(11)=   "Label4"
         Tab(0).Control(12)=   "Label3"
         Tab(0).Control(13)=   "Label2"
         Tab(0).Control(14)=   "Label1"
         Tab(0).ControlCount=   15
         TabCaption(1)   =   "Modificar"
         TabPicture(1)   =   "frmnuevooperario.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdcancelarmod"
         Tab(1).Control(1)=   "cmdmodificar"
         Tab(1).Control(2)=   "RichTextBox2"
         Tab(1).Control(3)=   "txtdireccionmod"
         Tab(1).Control(4)=   "txttelefonomod"
         Tab(1).Control(5)=   "txtcedulamod"
         Tab(1).Control(6)=   "txtapellidosmod"
         Tab(1).Control(7)=   "txtnombremod"
         Tab(1).Control(8)=   "cbonombre"
         Tab(1).Control(9)=   "Label12"
         Tab(1).Control(10)=   "Label11"
         Tab(1).Control(11)=   "Label10"
         Tab(1).Control(12)=   "Label9"
         Tab(1).Control(13)=   "Label8"
         Tab(1).Control(14)=   "Label7"
         Tab(1).ControlCount=   15
         TabCaption(2)   =   "Eliminar"
         TabPicture(2)   =   "frmnuevooperario.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmdeliminar"
         Tab(2).Control(1)=   "cmdcancelareli"
         Tab(2).Control(2)=   "RichTextBox3"
         Tab(2).Control(3)=   "cbonombreeli"
         Tab(2).Control(4)=   "Label13"
         Tab(2).ControlCount=   5
         TabCaption(3)   =   "Ver"
         TabPicture(3)   =   "frmnuevooperario.frx":0054
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "fleoperario"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).ControlCount=   1
         Begin GradientCommand.GGCommand cmdcancelareli 
            Height          =   495
            Left            =   -72120
            TabIndex        =   36
            Top             =   2520
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
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
         Begin GradientCommand.GGCommand cmdeliminar 
            Height          =   495
            Left            =   -74400
            TabIndex        =   35
            Top             =   2520
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
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
         Begin GradientCommand.GGCommand cmdguardar 
            Height          =   495
            Left            =   -74640
            TabIndex        =   12
            Top             =   3720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
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
         Begin GradientCommand.GGCommand cmdmodificar 
            Height          =   495
            Left            =   -74640
            TabIndex        =   31
            Top             =   3720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
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
            Height          =   495
            Left            =   -72240
            TabIndex        =   13
            Top             =   3720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   873
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
         Begin GradientCommand.GGCommand cmdcancelarmod 
            Height          =   495
            Left            =   -72120
            TabIndex        =   32
            Top             =   3720
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   873
            Caption         =   "Limpiar"
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
         Begin RichTextLib.RichTextBox RichTextBox3 
            Height          =   3855
            Left            =   -70440
            TabIndex        =   37
            Top             =   480
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   6800
            _Version        =   393217
            BackColor       =   -2147483633
            ReadOnly        =   -1  'True
            TextRTF         =   $"frmnuevooperario.frx":0070
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Proyecto1.SComboBox cbonombreeli 
            Height          =   450
            Left            =   -72840
            TabIndex        =   33
            Top             =   1200
            Width           =   1935
            _ExtentX        =   3413
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
         Begin RichTextLib.RichTextBox RichTextBox2 
            Height          =   3855
            Left            =   -70440
            TabIndex        =   30
            Top             =   480
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   6800
            _Version        =   393217
            BackColor       =   -2147483633
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            TextRTF         =   $"frmnuevooperario.frx":0138
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox txtdireccionmod 
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
            Left            =   -72840
            TabIndex        =   29
            Top             =   3120
            Width           =   1935
         End
         Begin VB.TextBox txttelefonomod 
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
            Left            =   -72840
            TabIndex        =   28
            Top             =   2640
            Width           =   1935
         End
         Begin VB.TextBox txtcedulamod 
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
            Left            =   -72840
            TabIndex        =   27
            Top             =   2160
            Width           =   1935
         End
         Begin VB.TextBox txtapellidosmod 
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
            Left            =   -72840
            TabIndex        =   26
            Top             =   1680
            Width           =   1935
         End
         Begin VB.TextBox txtnombremod 
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
            Left            =   -72840
            TabIndex        =   25
            Top             =   1200
            Width           =   1935
         End
         Begin Proyecto1.SComboBox cbonombre 
            Height          =   450
            Left            =   -72840
            TabIndex        =   18
            Top             =   720
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
         Begin ubGridControl.ubGrid fleoperario 
            Height          =   4095
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   6975
            _ExtentX        =   12303
            _ExtentY        =   7223
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
         Begin VB.TextBox txtapellidos 
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
            Left            =   -72840
            TabIndex        =   16
            Top             =   1680
            Width           =   1935
         End
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   3855
            Left            =   -70440
            TabIndex        =   14
            Top             =   480
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   6800
            _Version        =   393217
            BackColor       =   -2147483633
            ReadOnly        =   -1  'True
            TextRTF         =   $"frmnuevooperario.frx":0206
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.TextBox txtdireccion 
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
            Left            =   -72840
            TabIndex        =   6
            Top             =   3120
            Width           =   1935
         End
         Begin VB.TextBox txttelefono 
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
            Left            =   -72840
            TabIndex        =   5
            Top             =   2640
            Width           =   1935
         End
         Begin VB.TextBox txtcedula 
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
            Left            =   -72840
            TabIndex        =   4
            Top             =   2160
            Width           =   1935
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
            Left            =   -72840
            TabIndex        =   3
            Top             =   1200
            Width           =   1935
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
            Left            =   -72840
            TabIndex        =   2
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label13 
            Caption         =   "Nombre:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   -74400
            TabIndex        =   34
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label12 
            Caption         =   "Nuevo Telefono:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -75000
            TabIndex        =   24
            Top             =   2760
            Width           =   2175
         End
         Begin VB.Label Label11 
            Caption         =   "Nueva Cedula:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -75000
            TabIndex        =   23
            Top             =   2280
            Width           =   2055
         End
         Begin VB.Label Label10 
            Caption         =   "Nuevos Apellidos:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -75000
            TabIndex        =   22
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label Label9 
            Caption         =   "Nueva Dirección:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -75000
            TabIndex        =   21
            Top             =   3240
            Width           =   2175
         End
         Begin VB.Label Label8 
            Caption         =   "Nuevo Nombre:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -75000
            TabIndex        =   20
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label Label7 
            Caption         =   "Nombre:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   -75000
            TabIndex        =   19
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label Label6 
            Caption         =   "Apellidos:"
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
            Left            =   -74640
            TabIndex        =   15
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Dirección:"
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
            Left            =   -74640
            TabIndex        =   11
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Telefono:"
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
            Left            =   -74640
            TabIndex        =   10
            Top             =   2640
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Cedula:"
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
            Left            =   -74640
            TabIndex        =   9
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Nombre:"
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
            Left            =   -74640
            TabIndex        =   8
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label Label1 
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
            Height          =   375
            Left            =   -74640
            TabIndex        =   7
            Top             =   720
            Width           =   1215
         End
      End
   End
   Begin HookMenu.XpMenu XpMenu1 
      Left            =   0
      Top             =   0
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   3
      CheckBorderColor=   7021576
      SelMenuBorder   =   0
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
      Bmp:1           =   "frmnuevooperario.frx":02DA
      Key:1           =   "#mnaarchivosalir"
      Bmp:2           =   "frmnuevooperario.frx":0702
      Key:2           =   "#mnuayudaayuda"
      Bmp:3           =   "frmnuevooperario.frx":0B2A
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
   Begin VB.Menu mnuayuda 
      Caption         =   "?"
      Begin VB.Menu mnuayudaayuda 
         Caption         =   "Ayuda"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmnuevooperario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cargartxt()
cerrarconsulta
sentencia = "select * from operario where nombre='" & cbonombre.Text & "'"
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic
If Not consulta.RecordCount = 0 Then
txtnombremod.Text = consulta.Fields("nombre")
txtapellidosmod.Text = consulta.Fields("apellidos")
txtcedulamod.Text = consulta.Fields("cedula")
txttelefonomod.Text = consulta.Fields("telefono")
txtdireccionmod.Text = consulta.Fields("direccion")
End If
End Sub
Private Sub cargarnombre()
cbonombre.Clear
cbonombreeli.Clear
cerrarconsulta
consultar "select nombre from operario"
If Not consulta.RecordCount = 0 Then
Do Until consulta.EOF
 cbonombre.AddItem (consulta.Fields("nombre"))
 cbonombreeli.AddItem (consulta.Fields("nombre"))
  consulta.MoveNext
Loop
consulta.MoveFirst
cbonombre.Text = consulta.Fields("nombre")
End If
End Sub

Private Sub cbonombre_SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
cargartxt
End Sub

Private Sub cmdcancelar_Click()
txtnombre.Text = ""
txtapellidos.Text = ""
txtcedula.Text = ""
txttelefono.Text = ""
txtdireccion.Text = ""
cargarcodigos
End Sub

Private Sub cmdcancelareli_Click()
cbonombre.ListIndex = -1
End Sub

Private Sub cmdcancelarmod_Click()
txtnombremod.Text = ""
txtapellidosmod.Text = ""
txtcedulamod.Text = ""
txttelefonomod.Text = ""
txtdireccionmod.Text = ""
End Sub

Private Sub cmdeliminar_Click()
If MsgBox("Desea eliminar el operario?", vbQuestion + vbYesNo, "Eliminar") = vbYes Then
 If cbonombreeli.Text = "" Then
  MsgBox ("Elija un nombre"), vbExclamation, "Escojer"
  Else
   sentencia = "delete from operario where nombre='" & cbonombreeli.Text & "'"
   conexion.Execute sentencia
   cbonombre.Clear
   cbonombreeli.Clear
   cargarnombre
   cargaroperario
   txtnombremod.Text = ""
   txtapellidosmod.Text = ""
   txtcedulamod.Text = ""
   txttelefonomod.Text = ""
   txtdireccionmod.Text = ""
 End If
End If
End Sub

Private Sub cmdguardar_Click()
If txtnombre.Text = "" Then
 MsgBox ("Escriba un nombre"), vbExclamation, "Escribir"
 Else
  If txtapellidos.Text = "" Then
   MsgBox ("Escriba los apellidos"), vbExclamation, "Escribir"
   Else
    If txtcedula.Text = "" Then
     MsgBox ("Escriba un numero de cedula"), vbExclamation, "Escribir"
     Else
      If txttelefono.Text = "" Then
       MsgBox ("Escriba un telefono"), vbExclamation, "Escribir"
       Else
        If txtdireccion.Text = "" Then
         MsgBox ("Escriba una direccion"), vbExclamation, "Escribir"
         Else
          If MsgBox("Desea guardar este operario?", vbQuestion + vbYesNo, "Guardar") = vbYes Then
           operario.AddNew
           operario.Fields("codigo") = txtcodigo.Text
           operario.Fields("nombre") = txtnombre.Text
           operario.Fields("apellidos") = txtapellidos.Text
           operario.Fields("cedula") = txtcedula.Text
           operario.Fields("telefono") = txttelefono.Text
           operario.Fields("direccion") = txtdireccion.Text
           operario.Update
           MsgBox ("El operario fue guardado exitosamente"), vbInformation, "Guardado"
           txtnombre.Text = ""
           txtapellidos.Text = ""
           txtcedula.Text = ""
           txttelefono.Text = ""
           txtdireccion.Text = ""
           cargarcodigos
           cargaroperario
          End If
        End If
      End If
    End If
  End If
End If
End Sub
Private Sub cmdmodificar_Click()
If txtnombremod.Text = "" Then
 MsgBox ("Escriba un nombre"), vbExclamation, "Escribir"
 Else
  If txtapellidosmod.Text = "" Then
   MsgBox ("Escriba los apellidos"), vbExclamation, "Escribir"
   Else
    If txtcedulamod.Text = "" Then
     MsgBox ("Escriba un numero de cedula"), vbExclamation, "Escribir"
     Else
      If txttelefonomod.Text = "" Then
       MsgBox ("Escriba un telefono"), vbExclamation, "Escribir"
       Else
        If txtdireccionmod.Text = "" Then
         MsgBox ("Escriba una direccion"), vbExclamation, "Escribir"
         Else
          If MsgBox("Desea modificar este operario?", vbQuestion + vbYesNo, "Modificar") = vbYes Then
           sentencia = "update operario set nombre='" & txtnombremod.Text & "'" & ", apellidos='" & txtapellidosmod.Text & "'" & ", cedula='" & txtcedulamod.Text & "'" & ", telefono='" & txttelefonomod.Text & "'" & ", direccion='" & txtdireccionmod.Text & "'" & " where nombre='" & cbonombre.Text & "'"
           conexion.Execute sentencia
           MsgBox ("El operario fue modificado exitosamente"), vbInformation, "Exito"
           txtnombremod.Text = ""
           txtapellidosmod.Text = ""
           txtcedulamod.Text = ""
           txttelefonomod.Text = ""
           txtdireccionmod.Text = ""
           cbonombre.Clear
           cbonombreeli.Clear
           cargarnombre
           cargaroperario
          End If
        End If
      End If
    End If
  End If
End If
End Sub

Private Sub Form_Load()
abrirtablaoperario
cargarcodigos
cargarnombre
cargartxt
cargaroperario
autoajustar
cargaricono Me
End Sub
Private Sub cargarcodigos()
cargarcodigo "operario", Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
If operario.State = 1 Then
 operario.Close
End If
End Sub

Private Sub mnuarchivosalir_Click()
Unload Me
End Sub

Private Sub sstoperario_Click(PreviousTab As Integer)
If sstoperario.Caption = "Modificar" Then
cbonombre.Clear
cbonombreeli.Clear
 cargarnombre
 cargartxt
End If
If sstoperario.Caption = "Eliminar" Then
 cbonombreeli.Clear
 cbonombre.Clear
 cargarnombre
End If
If sstoperario.Caption = "ver" Then
 cargaroperario
 autoajustar
End If
If sstoperario.Caption = "Nuevo" Then
cargarcodigo "operario", Me
End If
End Sub
Private Sub txtapellidos_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Or keyascii = 241 Or keyascii = 209 Then
 If keyascii = 13 Then
  txtcedula.SetFocus
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo letras y numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub txtapellidosmod_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Or keyascii = 241 Or keyascii = 209 Then
 If keyascii = 13 Then
  txtcedulamod.SetFocus
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo letras y numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub txtcedulamod_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Then
 If keyascii = 13 Then
  txttelefonomod.SetFocus
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub

Private Sub txtcedula_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Then
 If keyascii = 13 Then
  txttelefono.SetFocus
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub txtdireccionmod_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Or keyascii = 45 Or keyascii = 127 Then
 If keyascii = 13 Then
  cmdmodificar_Click
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo letras y numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub txtdireccion_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Or keyascii = 45 Or keyascii = 127 Then
 If keyascii = 13 Then
  cmdguardar_Click
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo letras y numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub txtnombremod_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Or keyascii = 241 Or keyascii = 209 Then
 If keyascii = 13 Then
  txtapellidosmod.SetFocus
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo letras y numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub txtnombre_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Or keyascii = 241 Or keyascii = 209 Then
 If keyascii = 13 Then
  txtapellidos.SetFocus
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo letras y numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub txttelefono_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Then
 If keyascii = 13 Then
  txtdireccion.SetFocus
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub txttelefonomod_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Then
 If keyascii = 13 Then
  txtdireccionmod.SetFocus
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub cargaroperario()
Dim i As Integer
cerrarconsulta
sentencia = "Select * from operario"
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic

fleoperario.AutoSetup consulta.RecordCount + 1, 6, True, True, "Codigo |Nombre |Apellidos |Cedula |Telefono |Dirección"

If Not consulta.RecordCount = 0 Then
i = 1
Do Until consulta.EOF
 fleoperario.Row = i
 fleoperario.Col = 0
 fleoperario.Text = i
 
 fleoperario.Col = 1
 fleoperario.Text = consulta.Fields("codigo")
 
 fleoperario.Col = 2
 fleoperario.Text = consulta.Fields("nombre")
 
 fleoperario.Col = 3
 fleoperario.Text = consulta.Fields("apellidos")
 
 fleoperario.Col = 4
 fleoperario.Text = consulta.Fields("cedula")
 
 fleoperario.Col = 5
 fleoperario.Text = consulta.Fields("telefono")
 
 fleoperario.Col = 6
 fleoperario.Text = consulta.Fields("direccion")
 
 i = i + 1
 consulta.MoveNext
Loop
End If

End Sub
Private Sub autoajustar()
With fleoperario
.ColWidth(1) = 50
.ColWidth(2) = 100
.ColWidth(3) = 200
.ColWidth(4) = 70
.ColWidth(5) = 70
.ColWidth(6) = 100
End With
End Sub
