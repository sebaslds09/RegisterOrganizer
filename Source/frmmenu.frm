VERSION 5.00
Object = "{4EC837EB-4201-4C6D-A7D7-0C99D69CD9A6}#1.0#0"; "GradientCommand.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Object = "{7E02BD11-35B7-472C-8187-1538973C2A67}#32.0#0"; "NotifyIcon.ocx"
Begin VB.Form frmmenu 
   BackColor       =   &H00EECB82&
   Caption         =   "Menu"
   ClientHeight    =   10200
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   19080
   Icon            =   "frmmenu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10200
   ScaleWidth      =   19080
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin GradientCommand.GGCommand cmdnuevo 
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   5160
      Width           =   2055
      _ExtentX        =   3625
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
   Begin GradientCommand.GGCommand cmdproduccion 
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   4200
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      Caption         =   "Producción"
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
   Begin GradientCommand.GGCommand cmdmateria 
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   3240
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      Caption         =   "Mateia Prima"
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
   Begin GradientCommand.GGCommand cmdvereportes 
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   2280
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      Caption         =   "Ver Reportes"
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
   Begin GradientCommand.GGCommand cmdreporte 
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1085
      Caption         =   "Nuevo Reporte"
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
   Begin NotifyIcon.NotifyIconCtrl ntc 
      Left            =   960
      Top             =   360
      _ExtentX        =   1508
      _ExtentY        =   714
      Icon            =   "frmmenu.frx":0CCA
      ToolTip         =   "Register Organizer"
      BalloonText     =   "Register Organizer Ejecutandose"
      BalloonTitle    =   "Register Organizer"
      BalloonIcon     =   1
   End
   Begin HookMenu.XpMenu XpMenu1 
      Left            =   240
      Top             =   240
      _ExtentX        =   900
      _ExtentY        =   900
      BitmapSize      =   20
      BmpCount        =   14
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
      Bmp:1           =   "frmmenu.frx":19A4
      Key:1           =   "#mnuarchivonuevo"
      Bmp:2           =   "frmmenu.frx":1DCC
      Key:2           =   "#mnuarchivosalir"
      Bmp:3           =   "frmmenu.frx":21F4
      Key:3           =   "#mnunuevooperario"
      Bmp:4           =   "frmmenu.frx":261C
      Key:4           =   "#mnunuevocliente"
      Bmp:5           =   "frmmenu.frx":2A44
      Key:5           =   "#mnunuevohilo"
      Bmp:6           =   "frmmenu.frx":2E6C
      Key:6           =   "#mnunuevomaquina"
      Bmp:7           =   "frmmenu.frx":3294
      Key:7           =   "#mnunuevotipomaquina"
      Bmp:8           =   "frmmenu.frx":36BC
      Key:8           =   "#mnunuevocausa"
      Bmp:9           =   "frmmenu.frx":3AE4
      Key:9           =   "#mnunuevotiempo"
      Bmp:10          =   "frmmenu.frx":3F0C
      Key:10          =   "#mnumateriainventario"
      Bmp:11          =   "frmmenu.frx":4334
      Key:11          =   "#mnuproduccioncrear"
      Bmp:12          =   "frmmenu.frx":475C
      Key:12          =   "#mnureportegenerar"
      Bmp:13          =   "frmmenu.frx":4B84
      Key:13          =   "#mnureportesver"
      Bmp:14          =   "frmmenu.frx":4FAC
      Key:14          =   "#mnuarchivocerrar"
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
   Begin VB.Timer hora 
      Interval        =   1000
      Left            =   0
      Top             =   9480
   End
   Begin MSComctlLib.StatusBar stb1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9825
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2646
            Text            =   "Bordados Marion"
            TextSave        =   "Bordados Marion"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   10800
      Left            =   2760
      Picture         =   "frmmenu.frx":53D4
      Top             =   -360
      Width           =   14400
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuarchivonuevo 
         Caption         =   "Nuevo Usuario"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuarchivocerrar 
         Caption         =   "Cerrar Sesión"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuarchivosalir 
         Caption         =   "Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnunuevo 
      Caption         =   "Nuevo"
      Begin VB.Menu mnunuevooperario 
         Caption         =   "Operario"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnunuevocliente 
         Caption         =   "Cliente"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnunuevohilo 
         Caption         =   "Hilo"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnunuevomaquina 
         Caption         =   "Maquina"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnunuevotipomaquina 
         Caption         =   "Tipo Maquina"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnunuevocausa 
         Caption         =   "Causa de Tiempo"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnunuevotiempo 
         Caption         =   "Tiempo Muerto"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnumateria 
      Caption         =   "Materia"
      Begin VB.Menu mnumateriainventario 
         Caption         =   "Inventario de Materia"
         Shortcut        =   ^I
      End
   End
   Begin VB.Menu mnuproduccion 
      Caption         =   "Producción"
      Begin VB.Menu mnuproduccioncrear 
         Caption         =   "Produccion"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnureportes 
      Caption         =   "Reportes"
      Begin VB.Menu mnureportegenerar 
         Caption         =   "Generar Reporte"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnureportesver 
         Caption         =   "Ver Reportes"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mnuayuda 
      Caption         =   "?"
      Begin VB.Menu mnuayudaayuda 
         Caption         =   "Ayuda"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuayudair 
         Caption         =   "Ir a la Pagina web del Autor"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu mnutray 
      Caption         =   "tray"
      Visible         =   0   'False
      Begin VB.Menu mnutrayrestaurar 
         Caption         =   "Restaurar"
      End
      Begin VB.Menu mnutraysalir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "frmmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdmateria_Click()
frmmateria.Show vbModal
End Sub

Private Sub cmdnuevo_Click()
frmnuevo.Show vbModal
End Sub

Private Sub cmdproduccion_Click()
frmproduccion.Show vbModal
End Sub

Private Sub cmdreportes_Click()
frmreportes.Show vbModal
End Sub

Private Sub cmdreporte_Click()
frmreportes.Show vbModal
End Sub

Private Sub cmdvereportes_Click()
frmverreportes.Show vbModal
End Sub

Private Sub Form_Load()
Esperar 1
hora_Timer
setntc
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then
        ntc.ShowBalloon
        Me.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Usuario.State = 1 Then
 Usuario.Close
End If
Unload frmusuario
If hilo.State = 1 Then
 hilo.Close
End If
If cliente.State = 1 Then
 cliente.Close
End If
If maquina.State = 1 Then
 maquina.Close
End If
If tipo_maquina.State = 1 Then
 tipo_maquina.Close
End If
If causa.State = 1 Then
 causa.Close
End If
If produccion.State = 1 Then
 produccion.Close
End If
If tiempo_muerto.State = 1 Then
 tiempo_muerto.Close
End If
If materia_prima.State = 1 Then
 materia_prima.Close
End If
If operario.State = 1 Then
 operario.Close
End If
If reporte.State = 1 Then
 reporte.Close
End If
If reporte_detalle.State = 1 Then
 reporte_detalle.Close
End If
If suiche.State = 1 Then
 suiche.Close
End If
If produccion_hilo.State = 1 Then
 produccion_hilo.Close
End If
Unload frmsuiche
cerrarconsulta
cerrarconsulta2
End Sub

Private Sub mnuarchivocerrar_Click()
Unload Me
frmusuario.Show
End Sub

Private Sub mnuarchivonuevo_Click()
st = 1
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
frmusuario.Show vbModal
End Sub

Private Sub mnuarchivosalir_Click()
Unload Me
End Sub

Private Sub mnuayudair_Click()
Call Shell("C:\Archiv~1\Internet Explorer\IEXPLORE.EXE www.10-5lomejor.tk")
End Sub

Private Sub mnumateriainventario_Click()
frmmateria.Show vbModal
End Sub

Private Sub mnunuevocausa_Click()
frmcausa.Show vbModal
End Sub

Private Sub mnunuevocliente_Click()
frmnuevocliente.Show vbModal
End Sub

Private Sub mnunuevohilo_Click()
frmhilo.Show vbModal
End Sub

Private Sub mnunuevomaquina_Click()
frmmaquina.Show vbModal
End Sub

Private Sub mnunuevooperario_Click()
frmnuevooperario.Show vbModal
End Sub

Private Sub mnuvercliente_Click()
frmnuevocliente.Show vbModal
End Sub
Private Sub hora_Timer()
stb1.Panels(2).Text = time
End Sub

Private Sub mnunuevotiempo_Click()
frmtiempo.Show vbModal
End Sub

Private Sub mnunuevotipomaquina_Click()
frmtipomaquina.Show vbModal
End Sub

Private Sub mnuproduccioncrear_Click()
frmproduccion.Show vbModal
End Sub

Private Sub mnureportegenerar_Click()
frmreportes.Show vbModal
End Sub

Private Sub mnureportesver_Click()
frmverreportes.Show vbModal
End Sub
Private Sub setntc()
ntc.Visible = True
End Sub
Private Sub mnutrayrestaurar_Click()
    Me.WindowState = 2
    Me.Show
End Sub

Private Sub mnutraysalir_Click()
    If MsgBox(" Salir ?? ", vbQuestion + vbYesNo) = vbYes Then
        End
    End If
End Sub
Private Sub ntc_ContextMenu()
    PopupMenu mnutray
End Sub
Private Sub ntc_DblClick(Button As Integer)
    If Button = 1 Then
       mnutrayrestaurar_Click
    End If
End Sub
