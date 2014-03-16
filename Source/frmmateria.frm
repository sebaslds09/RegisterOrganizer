VERSION 5.00
Object = "{4EC837EB-4201-4C6D-A7D7-0C99D69CD9A6}#1.0#0"; "GradientCommand.ocx"
Object = "{13592B48-653C-491D-ACB1-C3140AA12F33}#6.1#0"; "ubGrid.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Begin VB.Form frmmateria 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Materia Prima"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   7170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin HookMenu.XpMenu XpMenu1 
      Left            =   6240
      Top             =   6240
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   2
      CheckBorderColor=   7021576
      SelMenuBorder   =   7021576
      SelMenuBackColor=   14073525
      SelMenuForeColor=   12582912
      SelCheckBackColor=   13740436
      MenuBorderColor =   6956042
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
      Bmp:1           =   "frmmateria.frx":0000
      Key:1           =   "#mnuarchivosalir"
      Bmp:2           =   "frmmateria.frx":0428
      Key:2           =   "#mnuayudaayuda"
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
      Height          =   8055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14295
      _ExtentX        =   25215
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
      Begin GradientCommand.GGCommand cmdadicionar 
         Height          =   495
         Left            =   5640
         TabIndex        =   8
         Top             =   1080
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Caption         =   "Adicionar"
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
         ToolTip         =   "Adicionar"
         ToolTipTitle    =   "Bordados Marion"
         ToolTipForeColor=   16711680
         ToolTipIcon     =   1
         ToolTipStyle    =   1
         Border3DStyle   =   2
         GradientToColor =   15649666
      End
      Begin GradientCommand.GGCommand cmdmodificar 
         Height          =   495
         Left            =   5640
         TabIndex        =   9
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
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
         ToolTip         =   "Modificar"
         ToolTipTitle    =   "Bordados Marion"
         ToolTipForeColor=   16711680
         ToolTipIcon     =   1
         ToolTipStyle    =   1
         Border3DStyle   =   2
         GradientToColor =   15649666
      End
      Begin GradientCommand.GGCommand cmdeliminar 
         Height          =   495
         Left            =   5640
         TabIndex        =   10
         Top             =   2520
         Width           =   1335
         _ExtentX        =   2355
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
         ToolTip         =   "Eliminar"
         ToolTipTitle    =   "Bordados Marion"
         ToolTipForeColor=   16711680
         ToolTipIcon     =   1
         ToolTipStyle    =   1
         Border3DStyle   =   2
         GradientToColor =   15649666
      End
      Begin GradientCommand.GGCommand cmdguardar 
         Height          =   495
         Left            =   5640
         TabIndex        =   11
         Top             =   3240
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
         ToolTip         =   "Guardar"
         ToolTipTitle    =   "Bordados Marion"
         ToolTipForeColor=   16711680
         ToolTipIcon     =   1
         ToolTipStyle    =   1
         Border3DStyle   =   2
         GradientToColor =   15649666
      End
      Begin GradientCommand.GGCommand cmdcancelar 
         Height          =   495
         Left            =   5640
         TabIndex        =   12
         Top             =   3960
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
         ToolTip         =   "Cancelar"
         ToolTipTitle    =   "Bordados Marion"
         ToolTipForeColor=   16711680
         ToolTipIcon     =   1
         ToolTipStyle    =   1
         Border3DStyle   =   2
         GradientToColor =   15649666
      End
      Begin GradientCommand.GGCommand cmdbuscar 
         Height          =   495
         Left            =   5640
         TabIndex        =   13
         Top             =   4680
         Width           =   1335
         _ExtentX        =   2355
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
         ToolTip         =   "Buscar"
         ToolTipTitle    =   "Bordados Marion"
         ToolTipForeColor=   16711680
         ToolTipIcon     =   1
         ToolTipStyle    =   1
         Border3DStyle   =   2
         GradientToColor =   15649666
      End
      Begin GradientCommand.GGCommand cmdsalir 
         Height          =   495
         Left            =   5640
         TabIndex        =   14
         Top             =   5400
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
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
         ToolTip         =   "Salir"
         ToolTipTitle    =   "Bordados Marion"
         ToolTipForeColor=   16711680
         ToolTipIcon     =   1
         ToolTipStyle    =   1
         Border3DStyle   =   2
         GradientToColor =   15649666
      End
      Begin ubGridControl.ubGrid flemateria 
         Height          =   4215
         Left            =   360
         TabIndex        =   7
         Top             =   2640
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   7435
         Rows            =   1
         Cols            =   3
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
         Left            =   2160
         TabIndex        =   6
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtmateria 
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
         Left            =   2160
         TabIndex        =   4
         Top             =   1200
         Width           =   2295
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
         Left            =   2160
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         Left            =   480
         TabIndex        =   5
         Top             =   1920
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00EECB82&
         Caption         =   "Materia:"
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
         Left            =   480
         TabIndex        =   3
         Top             =   1200
         Width           =   1065
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
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   960
      End
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
Attribute VB_Name = "frmmateria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rt As Integer
Private Sub cargarcodigos()
cargarcodigo "materia_prima", Me
End Sub
Private Sub cmdadicionar_Click()
cargarcodigos
habilitaraddmod
txtmateria.Text = ""
txtcantidad.Text = ""
rt = 1
End Sub
Private Sub cmdbuscar_Click()
frmbuscar.Show vbModal
End Sub
Private Sub cmdcancelar_Click()
habilitarguacan
      txtcodigo.Text = ""
      txtmateria.Text = ""
      txtcantidad.Text = ""
      cargarflemateria
      cargartxt
End Sub
Private Sub cmdeliminar_Click()
If MsgBox("Desea eliminar esta materia", vbQuestion + vbYesNo, "Eliminar") = vbYes Then
consultarotro "delete from materia_prima where codigo=" & txtcodigo.Text
      MsgBox ("La materia fue eliminada exitosamente"), vbInformation, "Guardar"
      txtcodigo.Text = ""
      txtmateria.Text = ""
      txtcantidad.Text = ""
      cargarflemateria
      cargartxt
    End If
End Sub
Private Sub cmdguardar_Click()
habilitarguacan
If rt = 1 Then
 If txtmateria.Text = "" Then
  MsgBox ("Escriba una materia prima"), vbInformation, "Escribir"
  Else
   If txtcantidad.Text = "" Then
    MsgBox ("Escriba una cantidad"), vbInformation, "Escibir"
     Else
      If MsgBox("Desea guardar esta materia", vbQuestion + vbYesNo, "Guardar") = vbYes Then
      materia_prima.AddNew
      materia_prima.Fields("codigo") = txtcodigo.Text
      materia_prima.Fields("materia") = txtmateria.Text
      materia_prima.Fields("cantidad") = txtcantidad.Text
      materia_prima.Update
      MsgBox ("La materia fue agregada exitosamente"), vbInformation, "Guardar"
      txtcodigo.Text = ""
      txtmateria.Text = ""
      txtcantidad.Text = ""
      cargarflemateria
      cargartxt
   End If
   End If
 End If
End If
If rt = 0 Then
 If txtmateria.Text = "" Then
  MsgBox ("Escriba una materia prima"), vbInformation, "Escribir"
  Else
   If txtcantidad.Text = "" Then
    MsgBox ("Escriba una cantidad"), vbInformation, "Escibir"
     Else
      consultarotro "update materia_prima set materia='" & txtmateria.Text & "'" & ", cantidad=" & txtcantidad.Text & " where codigo=" & txtcodigo.Text
      MsgBox ("La materia fue modificada exitosamente"), vbInformation, "Guardar"
      txtcodigo.Text = ""
      txtmateria.Text = ""
      txtcantidad.Text = ""
      cargarflemateria
      cargartxt
   End If
 End If
End If
End Sub
Private Sub cmdmodificar_Click()
habilitaraddmod
rt = 0
End Sub
Private Sub cmdsalir_Click()
Unload Me
End Sub
Private Sub flemateria_Click()
flemateria.SelectionMode = SelectionByRow
End Sub
Private Sub flemateria_DblClick()
If Not flemateria.Rows = 0 Then
txtcodigo.Text = flemateria.TextMatrix(flemateria.Row, 1)
txtmateria.Text = flemateria.TextMatrix(flemateria.Row, 2)
txtcantidad.Text = flemateria.TextMatrix(flemateria.Row, 3)
End If
End Sub
Private Sub Form_Load()
cargaricono Me
abrirtablamateria_prima
cargarflemateria
txtmateria.Enabled = False
txtcantidad.Enabled = False
cmdguardar.Enabled = False
cmdcancelar.Enabled = False
cargartxt
End Sub
Private Sub cargarflemateria()
Dim i As Integer
cerrarconsulta
sentencia = "select * from materia_prima"
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic
flemateria.AutoSetup consulta.RecordCount, 3, True, True, "Codigo |Materia |Cantidad "
flemateria.ColAllowEdit(1) = False
flemateria.ColAllowEdit(2) = False
flemateria.ColAllowEdit(3) = False

If Not consulta.RecordCount = 0 Then
i = i + 1
Do Until consulta.EOF
 With flemateria
 .Row = i
 .Col = 0
 .Text = i
 
 .Col = 1
 .Text = consulta.Fields("codigo")
 
 .Col = 2
 .Text = consulta.Fields("materia")
 
 .Col = 3
 .Text = consulta.Fields("cantidad")
 End With
 
 consulta.MoveNext
 i = i + 1
Loop
End If

End Sub
Private Sub mnuarchivosalir_Click()
Unload Me
End Sub
Private Sub autoajustar()
With flecliente
.ColWidth(1) = 50
.ColWidth(2) = 100
.ColWidth(3) = 70
End With
End Sub
Private Sub txtmateria_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Or keyascii = 241 Or keyascii = 209 Then
 If keyascii = 13 Then
 txtcantidad.SetFocus
keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo letras y numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub txtcantidad_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Then
 If keyascii = 13 Then
  cmdadicionar_Click
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub habilitaraddmod()
 txtmateria.Enabled = True
 txtcantidad.Enabled = True
 cmdadicionar.Enabled = False
 cmdmodificar.Enabled = False
 cmdeliminar.Enabled = False
 cmdguardar.Enabled = True
 cmdcancelar.Enabled = True
 cmdbuscar.Enabled = True
 cmdsalir.Enabled = False
End Sub
Private Sub habilitarguacan()
 txtmateria.Enabled = False
 txtcantidad.Enabled = False
 cmdadicionar.Enabled = True
 cmdmodificar.Enabled = True
 cmdeliminar.Enabled = True
 cmdguardar.Enabled = False
 cmdcancelar.Enabled = False
 cmdbuscar.Enabled = True
 cmdsalir.Enabled = True
End Sub
Private Sub cargartxt()
consultar "select * from materia_prima"

If Not consulta.RecordCount = 0 Then
consulta.MoveFirst
txtcodigo.Text = consulta.Fields("codigo")
txtmateria.Text = consulta.Fields("materia")
txtcantidad.Text = consulta.Fields("cantidad")
End If

End Sub
