VERSION 5.00
Object = "{4EC837EB-4201-4C6D-A7D7-0C99D69CD9A6}#1.0#0"; "GradientCommand.ocx"
Object = "{13592B48-653C-491D-ACB1-C3140AA12F33}#6.1#0"; "ubGrid.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Begin VB.Form frmnuevousuario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usuario"
   ClientHeight    =   5160
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   6960
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab sstusuario 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7005
      _ExtentX        =   12356
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
      TabPicture(0)   =   "frmnuevousuario.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdguardar"
      Tab(0).Control(1)=   "cmdcancelar"
      Tab(0).Control(2)=   "RichTextBox1"
      Tab(0).Control(3)=   "txtclave"
      Tab(0).Control(4)=   "txtrepclave"
      Tab(0).Control(5)=   "txtnombre"
      Tab(0).Control(6)=   "optusuario"
      Tab(0).Control(7)=   "optadministrador"
      Tab(0).Control(8)=   "Label3"
      Tab(0).Control(9)=   "Label2"
      Tab(0).Control(10)=   "Label1"
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Modificar"
      TabPicture(1)   =   "frmnuevousuario.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdguardarmod"
      Tab(1).Control(1)=   "cmdcancelarmod"
      Tab(1).Control(2)=   "RichTextBox2"
      Tab(1).Control(3)=   "txtrepclavemod"
      Tab(1).Control(4)=   "txtnombreclave"
      Tab(1).Control(5)=   "cbonombre"
      Tab(1).Control(6)=   "optclave"
      Tab(1).Control(7)=   "optnombre"
      Tab(1).Control(8)=   "lblrepclavemod"
      Tab(1).Control(9)=   "lblnombreclave"
      Tab(1).Control(10)=   "lblnombre"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Eliminar"
      TabPicture(2)   =   "frmnuevousuario.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdcancelareli"
      Tab(2).Control(1)=   "cmdeliminar"
      Tab(2).Control(2)=   "RichTextBox3"
      Tab(2).Control(3)=   "cbonombreeli"
      Tab(2).Control(4)=   "Label4"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Ver"
      TabPicture(3)   =   "frmnuevousuario.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "fleusuario"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin GradientCommand.GGCommand cmdcancelarmod 
         Height          =   495
         Left            =   -71760
         TabIndex        =   23
         Top             =   3120
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
      Begin GradientCommand.GGCommand cmdguardarmod 
         Height          =   495
         Left            =   -74520
         TabIndex        =   22
         Top             =   3120
         Width           =   1335
         _ExtentX        =   2355
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
         Left            =   -71760
         TabIndex        =   10
         ToolTipText     =   "Cancelar"
         Top             =   3120
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
      Begin GradientCommand.GGCommand cmdguardar 
         Height          =   495
         Left            =   -74520
         TabIndex        =   11
         ToolTipText     =   "Guardar"
         Top             =   3120
         Width           =   1335
         _ExtentX        =   2355
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
      Begin GradientCommand.GGCommand cmdeliminar 
         Height          =   495
         Left            =   -74400
         TabIndex        =   27
         Top             =   2760
         Width           =   1455
         _ExtentX        =   2566
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
      Begin GradientCommand.GGCommand cmdcancelareli 
         Height          =   495
         Left            =   -72000
         TabIndex        =   28
         Top             =   2760
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
      Begin RichTextLib.RichTextBox RichTextBox3 
         Height          =   3855
         Left            =   -70320
         TabIndex        =   29
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   6800
         _Version        =   393217
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         TextRTF         =   $"frmnuevousuario.frx":0070
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Proyecto1.SComboBox cbonombreeli 
         Height          =   435
         Left            =   -72600
         TabIndex        =   26
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   767
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
         ShadowColorText =   6582129
         Text            =   ""
      End
      Begin RichTextLib.RichTextBox RichTextBox2 
         Height          =   3855
         Left            =   -70320
         TabIndex        =   24
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   6800
         _Version        =   393217
         BackColor       =   -2147483633
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         TextRTF         =   $"frmnuevousuario.frx":0177
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txtrepclavemod 
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
         IMEMode         =   3  'DISABLE
         Left            =   -72600
         PasswordChar    =   "*"
         TabIndex        =   21
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox txtnombreclave 
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
         Left            =   -72600
         TabIndex        =   20
         Top             =   1800
         Width           =   2055
      End
      Begin Proyecto1.SComboBox cbonombre 
         Height          =   420
         Left            =   -72600
         TabIndex        =   16
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   741
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
         ShadowColorText =   6582129
         Text            =   ""
      End
      Begin VB.OptionButton optclave 
         Caption         =   "Clave"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -71880
         TabIndex        =   15
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optnombre 
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -74640
         TabIndex        =   14
         Top             =   600
         Width           =   1575
      End
      Begin ubGridControl.ubGrid fleusuario 
         Height          =   4095
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   7223
         Rows            =   1
         Cols            =   4
         Redraw          =   -1  'True
         ShowGrid        =   -1  'True
         GridSolid       =   -1  'True
         GridLineColor   =   12632256
         BorderStyle     =   2
         BackColorFixed  =   12632256
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
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
         AutoNewRow      =   -1  'True
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   3855
         Left            =   -70320
         TabIndex        =   12
         Top             =   480
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   6800
         _Version        =   393217
         BackColor       =   -2147483633
         TextRTF         =   $"frmnuevousuario.frx":02A2
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
      Begin VB.TextBox txtclave 
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
         IMEMode         =   3  'DISABLE
         Left            =   -72600
         PasswordChar    =   "*"
         TabIndex        =   9
         ToolTipText     =   "Clave"
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtrepclave 
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
         IMEMode         =   3  'DISABLE
         Left            =   -72600
         PasswordChar    =   "*"
         TabIndex        =   8
         ToolTipText     =   "Repetir Clave"
         Top             =   2280
         Width           =   2055
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
         Left            =   -72600
         TabIndex        =   7
         ToolTipText     =   "Nombre"
         Top             =   1320
         Width           =   2055
      End
      Begin VB.OptionButton optusuario 
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -71880
         TabIndex        =   3
         ToolTipText     =   "Usuario"
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optadministrador 
         Caption         =   "Administrador"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   -74640
         TabIndex        =   2
         ToolTipText     =   "Administrador"
         Top             =   600
         Width           =   2175
      End
      Begin VB.Label Label4 
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
         Left            =   -74160
         TabIndex        =   25
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblrepclavemod 
         Caption         =   "Repetir Clave:"
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
         Left            =   -74520
         TabIndex        =   19
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label lblnombreclave 
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
         Height          =   375
         Left            =   -74520
         TabIndex        =   18
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label lblnombre 
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
         Left            =   -74520
         TabIndex        =   17
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Repetir Clave:"
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
         Left            =   -74520
         TabIndex        =   6
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Clave:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74520
         TabIndex        =   5
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label1 
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
         Left            =   -74520
         TabIndex        =   4
         Top             =   1320
         Width           =   1095
      End
   End
   Begin XPFrame.FrameXp franuevousuario 
      Height          =   5175
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   9128
      BackColor       =   15649666
      Caption         =   "Bordados Marion"
      CaptionEstilo3D =   1
      BackColor       =   15649666
      ForeColor       =   12582912
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
      Bmp:1           =   "frmnuevousuario.frx":038B
      Key:1           =   "#mnaarchivosalir"
      Bmp:2           =   "frmnuevousuario.frx":07B3
      Key:2           =   "#mnuayudaayuda"
      Bmp:3           =   "frmnuevousuario.frx":0BDB
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
Attribute VB_Name = "frmnuevousuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As String

Private Sub cbonombre_SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
txtnombreclave.SetFocus
End Sub

Private Sub cmdcancelar_Click()
txtnombre.Text = ""
txtclave.Text = ""
txtrepclave.Text = ""
End Sub

Private Sub cmdeliminar_Click()
If MsgBox("Desea eliminar este usuario?", vbQuestion + vbYesNo, "Eliminar") = vbYes Then
If cbonombreeli.Text = "" Then
 MsgBox ("Selecione un nombre"), vbExclamation, "Selecionar"
  Else
  sentencia = "delete from usuario where nombre='" & cbonombreeli.Text & "'"
  conexion.Execute sentencia
  MsgBox ("El usuario fue eliminado exitosamente"), vbInformation, "Eliminado"
  cargarnombre
  cbonombreeli.ListIndex = -1
  cargarusuario
End If
End If
End Sub

Private Sub cmdguardar_Click()
If txtnombre.Text = "" Then
 MsgBox ("Ingrese un nombre"), vbExclamation, "Ingresar"
 Else
  If txtclave.Text = "" Then
   MsgBox ("Ingrese una clave"), vbExclamation, "Ingresar"
    Else
     If txtrepclave.Text = "" Then
      MsgBox ("Repita la clave"), vbExclamation, "Repetir"
       Else
        cerrarconsulta
        sentencia = "select nombre from usuario where nombre='" & txtnombre & "'"
        consulta.Source = sentencia
        consulta.Open , conexion, adOpenStatic
        If consulta.RecordCount = 0 Then
        If txtclave.Text = txtrepclave.Text Then
         Usuario.AddNew
         Usuario.Fields("nombre") = txtnombre.Text
         Usuario.Fields("clave") = txtclave.Text
         Usuario.Fields("tipo") = a
         Usuario.Update
         MsgBox ("El usuario fue guardado con exito"), vbInformation, "Exito"
         txtnombre.Text = ""
         txtclave.Text = ""
         txtrepclave.Text = ""
         cargarusuario
         Else
          MsgBox ("las claves no coinciden"), vbInformation, "Error"
          txtrepclave.Text = ""
        End If
        Else
        MsgBox ("El nombre ya existe"), vbExclamation, "Error"
        txtnombre.Text = ""
        txtclave.Text = ""
        txtrepclave.Text = ""
        End If
     End If
  End If
End If
End Sub

Private Sub cmdguardarmod_Click()
If optnombre.Value = True Then
 If txtnombreclave.Text = "" Then
 MsgBox ("Ingrese un nombre"), vbExclamation, "Ingresar"
 Else
 sentencia = "update usuario set nombre='" & txtnombreclave.Text & "'" & "where nombre='" & cbonombre.Text & "'"
 conexion.Execute sentencia
 MsgBox ("El nombre fue modificado exitosamente"), vbInformation, "Exito"
 txtnombreclave.Text = ""
 cbonombre.ListIndex = -1
 cargarnombre
 End If
End If
If optclave.Value = True Then
If txtnombreclave.Text = "" Then
 MsgBox ("Ingrese una clave"), vbExclamation, "Ingresar"
 Else
 If txtrepclavemod.Text = "" Then
 MsgBox ("Repita la clave"), vbExclamation, "Repetir"
 Else
 If txtnombreclave.Text = txtrepclavemod.Text Then
  sentencia = "update usuario set clave='" & txtrepclavemod.Text & "'" & "where nombre='" & cbonombre.Text & "'"
  conexion.Execute sentencia
  MsgBox ("La clave fue modificada exitosamente"), vbInformation, "Exito"
  txtnombreclave.Text = ""
  txtrepclavemod.Text = ""
  cbonombre.ListIndex = -1
  cargarusuario
  Else
   MsgBox ("Las claves no coinciden"), vbInformation, "Error"
   txtrepclavemod.Text = ""
 End If
 End If
 End If
End If
cargarusuario
End Sub

Private Sub Form_Load()
abrirtablausuario
Usuario.MoveFirst
optadministrador.Value = True
a = "Administrador"
optnombre.Value = True
habilitar
cargarnombre
cbonombre.ListIndex = -1
cargarusuario
cargaricono Me
End Sub

Private Sub mnuarchivosalir_Click()
Unload Me
End Sub

Private Sub optadministrador_Click()
 If optadministrador.Value = True Then
  a = "Administrador"
 End If
End Sub
Private Sub optclave_Click()
habilitar
End Sub

Private Sub optnombre_Click()
habilitar
End Sub

Private Sub optusuario_Click()
 If optusuario.Value = True Then
  a = "Usuario"
 End If
End Sub

Private Sub sstusuario_Click(PreviousTab As Integer)
If sstusuario.Caption = "Eliminar" Then
 cargarnombre
End If
If sstusuario.Caption = "Modificar" Then
 cargarnombre
End If

End Sub

Private Sub txtnombre_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Or keyascii = 241 Or keyascii = 209 Then
 If keyascii = 13 Then
  txtclave.SetFocus
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo letras y numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub txtclave_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Or keyascii = 241 Or keyascii = 209 Then
 If keyascii = 13 Then
  txtrepclave.SetFocus
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo letras y numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub txtnombreclave_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Or keyascii = 241 Or keyascii = 209 Then
 If keyascii = 13 Then
  If lblrepclavemod.Visible = True Then
  txtrepclavemod.SetFocus
  End If
  If lblrepclavemod.Visible = False Then
  cmdguardarmod_Click
  End If
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo letras y numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub txtrepclave_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Then
 If keyascii = 13 Then
  cmdguardar_Click
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo letras y numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub habilitar()
If optnombre.Value = True Then
 lblnombreclave.Caption = "Nuevo Nombre"
 lblrepclavemod.Visible = False
 txtrepclavemod.Visible = False
 txtnombreclave.PasswordChar = ""
End If
If optclave.Value = True Then
 lblnombreclave.Caption = "Clave"
 lblrepclavemod.Visible = True
 txtrepclavemod.Visible = True
 txtnombreclave.PasswordChar = "*"
End If
End Sub

Private Sub txtrepclavemod_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Then
 If keyascii = 13 Then
  cmdguardarmod_Click
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo letras y numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub cargarnombre()
Usuario.MoveFirst
cbonombre.Clear
cbonombreeli.Clear
abrirtablausuario

If Not Usuario.RecordCount = 0 Then
Do Until Usuario.EOF
 cbonombreeli.AddItem (Usuario.Fields("nombre"))
 cbonombre.AddItem (Usuario.Fields("nombre"))
 Usuario.MoveNext
Loop
End If

End Sub
Private Sub cargarusuario()
Dim i As Integer
cerrarconsulta
sentencia = "select * from usuario"
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic

fleusuario.AutoSetup consulta.RecordCount + 1, 3, True, True, "Nombre |Clave |Tipo"

If Not Usuario.RecordCount = 0 Then
i = 1
consulta.MoveFirst
Do Until consulta.EOF
 fleusuario.Row = i
 fleusuario.Col = 0
 fleusuario.Text = i
 
 fleusuario.Col = 1
 fleusuario.Text = consulta.Fields("nombre")
 
 fleusuario.Col = 2
 fleusuario.Text = consulta.Fields("clave")
 
 fleusuario.Col = 3
 fleusuario.Text = consulta.Fields("tipo")
 
 i = i + 1
 consulta.MoveNext
Loop
End If

End Sub
