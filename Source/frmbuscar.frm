VERSION 5.00
Object = "{4EC837EB-4201-4C6D-A7D7-0C99D69CD9A6}#1.0#0"; "GradientCommand.ocx"
Object = "{13592B48-653C-491D-ACB1-C3140AA12F33}#6.1#0"; "ubGrid.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Begin VB.Form frmbuscar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Buscar"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   8916
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
      Begin GradientCommand.GGCommand cmdbuscar 
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   873
         Caption         =   "Buscar"
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
         Left            =   1800
         TabIndex        =   4
         Top             =   1560
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
      Begin ubGridControl.ubGrid flenombre 
         Height          =   2055
         Left            =   480
         TabIndex        =   8
         Top             =   2880
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   3625
         Rows            =   1
         Cols            =   1
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
      Begin VB.PictureBox Picture2 
         Height          =   375
         Left            =   1200
         Picture         =   "frmbuscar.frx":0000
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   7
         Top             =   2280
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox Picture1 
         Height          =   375
         Left            =   600
         Picture         =   "frmbuscar.frx":058A
         ScaleHeight     =   315
         ScaleWidth      =   315
         TabIndex        =   6
         Top             =   2280
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00EECB82&
         Height          =   375
         Left            =   120
         Picture         =   "frmbuscar.frx":0B14
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2280
         Width           =   375
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
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00EECB82&
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmbuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdbuscar_Click()
consultar "select * from materia_prima where materia='" & txtnombre.Text & "'"
If Not consulta.RecordCount = 0 Then
 frmmateria.txtcodigo.Text = consulta.Fields("codigo")
 frmmateria.txtmateria.Text = consulta.Fields("materia")
 frmmateria.txtcantidad.Text = consulta.Fields("cantidad")
 Else
  If MsgBox("La materia prima no existe, desea agregarla", vbQuestion + vbYesNo, "Buscar") = vbYes Then
   frmmateria.cmdadicionar.Value = True
   frmmateria.txtmateria.Text = txtnombre.Text
  End If
End If
Unload Me
End Sub

Private Sub cmdcancelar_Click()
Unload Me
End Sub

Private Sub Command1_Click()
If Command1.Picture = Picture1.Picture Then
frmbuscar.Height = 5530
Command1.Picture = LoadPicture(App.Path & "\Logos\1008-archivos-formato-ico-de-16-pixeles-256-colores\arrow_up.ico")
Command1.Picture = Picture2.Picture
 Else
 If Command1.Picture = Picture2.Picture Then
frmbuscar.Height = 3285
Command1.Picture = LoadPicture(App.Path & "\Logos\1008-archivos-formato-ico-de-16-pixeles-256-colores\arrow_down.ico")
Command1.Picture = Picture1.Picture
End If
End If
End Sub

Private Sub Form_Load()
cargaricono Me
Command1.Picture = Picture1.Picture
cargarfle "select materia from materia_prima", "Nombre", "materia", 1, Me
End Sub

Private Sub txtnombre_Change()
cargarfle "select materia from materia_prima where materia like('" & txtnombre.Text & "%')" & " order by materia", "Nombre", "materia", 1, Me
End Sub
