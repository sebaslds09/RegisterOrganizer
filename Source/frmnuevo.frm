VERSION 5.00
Object = "{4EC837EB-4201-4C6D-A7D7-0C99D69CD9A6}#1.0#0"; "GradientCommand.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Begin VB.Form frmnuevo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nuevo"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin HookMenu.XpMenu XpMenu1 
      Left            =   4680
      Top             =   4080
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   9
      CheckBorderColor=   7021576
      SelMenuBorder   =   7021576
      SelMenuBackColor=   14073525
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
      Bmp:1           =   "frmnuevo.frx":0000
      Key:1           =   "#mnuarchivosalir"
      Bmp:2           =   "frmnuevo.frx":0428
      Key:2           =   "#mnuiroperarios"
      Bmp:3           =   "frmnuevo.frx":0850
      Key:3           =   "#mnuirclientes"
      Bmp:4           =   "frmnuevo.frx":0C78
      Key:4           =   "#mnuirhilos"
      Bmp:5           =   "frmnuevo.frx":10A0
      Key:5           =   "#mnuirmaquinas"
      Bmp:6           =   "frmnuevo.frx":14C8
      Key:6           =   "#mnuirtipo"
      Bmp:7           =   "frmnuevo.frx":18F0
      Key:7           =   "#mnuircausa"
      Bmp:8           =   "frmnuevo.frx":1D18
      Key:8           =   "#mnuirtiempo"
      Bmp:9           =   "frmnuevo.frx":2140
      Key:9           =   "#mnuayudaayuda"
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
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   8281
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
      Begin GradientCommand.GGCommand cmdtiempo 
         Height          =   615
         Left            =   3000
         TabIndex        =   7
         Top             =   2760
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   1085
         Caption         =   "Tiempo Muerto"
         BackColor       =   15649666
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
      Begin GradientCommand.GGCommand cmdcausa 
         Height          =   615
         Left            =   2880
         TabIndex        =   6
         Top             =   1680
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "Causa Tiempo"
         BackColor       =   15649666
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
      Begin GradientCommand.GGCommand cmdtipo 
         Height          =   615
         Left            =   2880
         TabIndex        =   5
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "Tipo Maquina"
         BackColor       =   15649666
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
      Begin GradientCommand.GGCommand cmdmaquinas 
         Height          =   615
         Left            =   360
         TabIndex        =   4
         Top             =   3720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "Maquinas"
         BackColor       =   15649666
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
      Begin GradientCommand.GGCommand cmdhilos 
         Height          =   615
         Left            =   360
         TabIndex        =   3
         Top             =   2760
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "Hilos"
         BackColor       =   15649666
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
      Begin GradientCommand.GGCommand cmdoperarios 
         Height          =   615
         Left            =   360
         TabIndex        =   1
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "Operarios"
         BackColor       =   15649666
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
      Begin GradientCommand.GGCommand cmdcientes 
         Height          =   615
         Left            =   360
         TabIndex        =   2
         Top             =   1680
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   1085
         Caption         =   "Clientes"
         BackColor       =   15649666
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
   Begin VB.Menu mnuarchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuarchivosalir 
         Caption         =   "Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuir 
      Caption         =   "&Ir"
      Begin VB.Menu mnuiroperarios 
         Caption         =   "Operarios"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuirclientes 
         Caption         =   "Clientes"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuirhilos 
         Caption         =   "Hilos"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuirmaquinas 
         Caption         =   "Maquinas"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuirtipo 
         Caption         =   "Tipo Maquinas"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuircausa 
         Caption         =   "Causa de Tiempo"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuirtiempo 
         Caption         =   "Tiempo Muerto"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuayuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu mnuayudaayuda 
         Caption         =   "Ayuda"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmnuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdcausa_Click()
frmcausa.Show vbModal
Unload Me
End Sub

Private Sub cmdcientes_Click()
frmnuevocliente.Show vbModal
Unload Me
End Sub

Private Sub cmdhilos_Click()
frmhilo.Show vbModal
Unload Me
End Sub

Private Sub cmdmaquinas_Click()
frmmaquina.Show vbModal
Unload Me
End Sub

Private Sub cmdoperarios_Click()
frmnuevooperario.Show vbModal
Unload Me
End Sub

Private Sub cmdtiempo_Click()
frmtiempo.Show vbModal
Unload Me
End Sub

Private Sub cmdtipo_Click()
frmtipomaquina.Show vbModal
Unload Me
End Sub


Private Sub mnuarchivosalir_Click()
Unload Me
End Sub

Private Sub mnuircausa_Click()
frmcausa.Show vbModal
Unload Me
End Sub

Private Sub mnuirclientes_Click()
frmnuevocliente.Show vbModal
Unload Me
End Sub

Private Sub mnuirhilos_Click()
frmhilo.Show vbModal
Unload Me
End Sub

Private Sub mnuirmaquinas_Click()
frmmaquina.Show vbModal
Unload Me
End Sub

Private Sub mnuiroperarios_Click()
frmnuevooperario.Show vbModal
Unload Me
End Sub

Private Sub mnuirtiempo_Click()
frmtiempo.Show vbModal
Unload Me
End Sub

Private Sub mnuirtipo_Click()
frmtipomaquina.Show vbModal
Unload Me
End Sub
