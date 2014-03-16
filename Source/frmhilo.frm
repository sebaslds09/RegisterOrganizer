VERSION 5.00
Object = "{4EC837EB-4201-4C6D-A7D7-0C99D69CD9A6}#1.0#0"; "GradientCommand.ocx"
Object = "{13592B48-653C-491D-ACB1-C3140AA12F33}#6.1#0"; "ubGrid.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Begin VB.Form frmhilo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hilo"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin HookMenu.XpMenu XpMenu1 
      Left            =   6720
      Top             =   0
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   2
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
      Bmp:1           =   "frmhilo.frx":0000
      Key:1           =   "#mnaarchivosalir"
      Bmp:2           =   "frmhilo.frx":0428
      Key:2           =   "#mnuayudaayuda"
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
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   9763
      BackColor       =   15649666
      Caption         =   "Bordados Marion"
      CaptionEstilo3D =   1
      BackColor       =   15649666
      ForeColor       =   16711680
      BordeColor      =   -2147483638
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
      Begin TabDlg.SSTab ssthilo 
         Height          =   4575
         Left            =   0
         TabIndex        =   1
         Top             =   480
         Width           =   7335
         _ExtentX        =   12938
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
         TabPicture(0)   =   "frmhilo.frx":0850
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "cmdguardar"
         Tab(0).Control(1)=   "cmdcancelar"
         Tab(0).Control(2)=   "RichTextBox1"
         Tab(0).Control(3)=   "txtcantidad"
         Tab(0).Control(4)=   "txtcolor"
         Tab(0).Control(5)=   "txtcodigo"
         Tab(0).Control(6)=   "Label3"
         Tab(0).Control(7)=   "Label2"
         Tab(0).Control(8)=   "Label1"
         Tab(0).ControlCount=   9
         TabCaption(1)   =   "Modificar"
         TabPicture(1)   =   "frmhilo.frx":086C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdcancelarmod"
         Tab(1).Control(1)=   "cmdguardarmod"
         Tab(1).Control(2)=   "RichTextBox2"
         Tab(1).Control(3)=   "txtnuevacantidad"
         Tab(1).Control(4)=   "txtnuevocolor"
         Tab(1).Control(5)=   "cbocolor"
         Tab(1).Control(6)=   "Label6"
         Tab(1).Control(7)=   "Label5"
         Tab(1).Control(8)=   "Label4"
         Tab(1).ControlCount=   9
         TabCaption(2)   =   "Eliminar"
         TabPicture(2)   =   "frmhilo.frx":0888
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmdcancelareli"
         Tab(2).Control(1)=   "cmdeliminar"
         Tab(2).Control(2)=   "RichTextBox3"
         Tab(2).Control(3)=   "cbocoloreli"
         Tab(2).Control(4)=   "Label7"
         Tab(2).ControlCount=   5
         TabCaption(3)   =   "Ver"
         TabPicture(3)   =   "frmhilo.frx":08A4
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "flehilo"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).ControlCount=   1
         Begin GradientCommand.GGCommand cmdcancelar 
            Height          =   495
            Left            =   -72480
            TabIndex        =   10
            Top             =   2880
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
         Begin GradientCommand.GGCommand cmdguardar 
            Height          =   495
            Left            =   -74640
            TabIndex        =   9
            Top             =   2880
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
         Begin GradientCommand.GGCommand cmdguardarmod 
            Height          =   495
            Left            =   -74520
            TabIndex        =   18
            Top             =   3120
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
         Begin GradientCommand.GGCommand cmdcancelarmod 
            Height          =   495
            Left            =   -72360
            TabIndex        =   19
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
         Begin GradientCommand.GGCommand cmdeliminar 
            Height          =   495
            Left            =   -74280
            TabIndex        =   21
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
         Begin GradientCommand.GGCommand cmdcancelareli 
            Height          =   495
            Left            =   -72240
            TabIndex        =   22
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
         Begin RichTextLib.RichTextBox RichTextBox3 
            Height          =   3855
            Left            =   -70320
            TabIndex        =   25
            Top             =   480
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   6800
            _Version        =   393217
            BackColor       =   -2147483633
            ReadOnly        =   -1  'True
            TextRTF         =   $"frmhilo.frx":08C0
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
            ReadOnly        =   -1  'True
            TextRTF         =   $"frmhilo.frx":098F
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
         Begin Proyecto1.SComboBox cbocoloreli 
            Height          =   450
            Left            =   -72720
            TabIndex        =   20
            Top             =   1200
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
         Begin VB.TextBox txtnuevacantidad 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   -72720
            TabIndex        =   17
            Top             =   2040
            Width           =   855
         End
         Begin VB.TextBox txtnuevocolor 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   -72720
            TabIndex        =   16
            Top             =   1560
            Width           =   1695
         End
         Begin Proyecto1.SComboBox cbocolor 
            Height          =   450
            Left            =   -72720
            TabIndex        =   12
            Top             =   1080
            Width           =   1695
            _ExtentX        =   2990
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
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   3855
            Left            =   -70320
            TabIndex        =   11
            Top             =   480
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   6800
            _Version        =   393217
            BackColor       =   -2147483633
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            TextRTF         =   $"frmhilo.frx":0A7C
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
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   -72960
            TabIndex        =   8
            Top             =   2040
            Width           =   735
         End
         Begin VB.TextBox txtcolor 
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   -72960
            TabIndex        =   6
            Top             =   1560
            Width           =   1695
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
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   -72960
            TabIndex        =   4
            Top             =   1080
            Width           =   975
         End
         Begin ubGridControl.ubGrid flehilo 
            Height          =   4095
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   7095
            _ExtentX        =   12515
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
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Color:"
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
            Left            =   -73680
            TabIndex        =   23
            Top             =   1200
            Width           =   780
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Nueva Cantidad:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   -74880
            TabIndex        =   15
            Top             =   2160
            Width           =   2055
         End
         Begin VB.Label Label5 
            Caption         =   "Nuevo Color:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   495
            Left            =   -74880
            TabIndex        =   14
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Color:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   -74880
            TabIndex        =   13
            Top             =   1080
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   -74640
            TabIndex        =   7
            Top             =   2160
            Width           =   1200
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Color:"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   -74640
            TabIndex        =   5
            Top             =   1680
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   -74640
            TabIndex        =   3
            Top             =   1200
            Width           =   960
         End
      End
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnaarchivosalir 
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
Attribute VB_Name = "frmhilo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cargarcodigos()
cargarcodigo "hilo", Me
End Sub

Private Sub cbocolor_SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
cargartxt
End Sub

Private Sub cmdcancelar_Click()
      txtcolor.Text = ""
      txtcantidad.Text = ""
End Sub

Private Sub cmdcancelarmod_Click()
     txtnuevocolor.Text = ""
     txtnuevacantidad.Text = ""
     cbocolor.Text = ""
End Sub

Private Sub cmdeliminar_Click()
If cbocoloreli.Text = "" Then
 MsgBox ("Elija un color"), vbExclamation, "Elejir"
 Else
 If MsgBox("Desea eliminar el hilo?", vbQuestion + vbYesNo, "Eliminar") = vbYes Then
  sentencia = "delete from hilo where color='" & cbocoloreli.Text & "'"
  conexion.Execute sentencia
  MsgBox ("El hilo fue eliminado exitosamente"), vbInformation, "Exito"
  cargarnombre
  cargarhilo
 End If
End If
End Sub

Private Sub cmdguardar_Click()
If txtcolor.Text = "" Then
 MsgBox ("Escriba un color"), vbExclamation, "Escribir"
  Else
   If txtcantidad.Text = "" Then
    MsgBox ("Escriba una cantidad"), vbExclamation, "Escribir"
     Else
     If MsgBox("Desea guardar este hilo", vbQuestion + vbYesNo, "Guardar") = vbYes Then
      hilo.AddNew
      hilo.Fields("codigo") = txtcodigo.Text
      hilo.Fields("color") = txtcolor.Text
      hilo.Fields("cantidad") = txtcantidad.Text
      hilo.Update
      MsgBox ("El hilo fue guardado con exito"), vbInformation, "Exito"
      txtcolor.Text = ""
      txtcantidad.Text = ""
      cargarnombre
      cargarhilo
      cargarcodigos
     End If
   End If
End If
End Sub

Private Sub cmdguardarmod_Click()
If txtnuevocolor.Text = "" Then
 MsgBox ("Escriba un color"), vbExclamation, "Escribir"
  Else
  If txtnuevacantidad.Text = "" Then
   MsgBox ("Escriba una cantidad"), vbExclamation, "Escribir"
    Else
    If MsgBox("Desea modificar este hilo?", vbQuestion + vbYesNo, "Modificar") = vbYes Then
     sentencia = "update hilo set color='" & txtnuevocolor.Text & "'" & ", cantidad=" & txtnuevacantidad.Text & " where color='" & cbocolor.Text & "'"
     conexion.Execute sentencia
     MsgBox ("El hilo fue modificado exitosamente"), vbInformation, "Exito"
     txtnuevocolor.Text = ""
     txtnuevacantidad.Text = ""
     cbocolor.Text = ""
     cargarnombre
     cargartxt
     cargarhilo
    End If
  End If
End If
End Sub

Private Sub Form_Load()
abrirtablahilo
cargarcodigos
cargarnombre
cargartxt
cargarhilo
cargaricono Me
End Sub

Private Sub mnaarchivosalir_Click()
Unload Me
End Sub

Private Sub ssthilo_Click(PreviousTab As Integer)
If ssthilo.Caption = "Modificar" Then
 cargarnombre
 cargartxt
End If
If ssthilo.Caption = "Nuevo" Then
 cargarcodigos
End If
If ssthilo.Caption = "Eliminar" Then
 cargarnombre
End If
If ssthilo.Caption = "Ver" Then
 cargarhilo
End If
End Sub
Private Sub txtcolor_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Or keyascii = 164 Or keyascii = 165 Then
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
  cmdguardar_Click
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub cargarnombre()
cerrarconsulta
sentencia = "select color from hilo"
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic
cbocolor.Clear
cbocoloreli.Clear

If Not consulta.RecordCount = 0 Then
Do Until consulta.EOF
 cbocolor.AddItem (consulta.Fields("color"))
 cbocoloreli.AddItem (consulta.Fields("color"))
 consulta.MoveNext
Loop
consulta.MoveFirst
cbocolor.Text = consulta.Fields("color")
End If

End Sub
Private Sub cargartxt()
cerrarconsulta
sentencia = "select * from hilo where color='" & cbocolor.Text & "'"
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic

If Not consulta.RecordCount = 0 Then
txtnuevocolor.Text = consulta.Fields("color")
txtnuevacantidad.Text = consulta.Fields("cantidad")
End If

End Sub
Private Sub txtnuevocolor_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Or keyascii = 164 Or keyascii = 165 Then
 If keyascii = 13 Then
  txtnuevacantidad.SetFocus
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo letras y numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub txtnuevacantidad_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Then
 If keyascii = 13 Then
  cmdguardarmod_Click
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub cargarhilo()
Dim i As Integer
cerrarconsulta
sentencia = "select * from hilo"
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic

flehilo.AutoSetup consulta.RecordCount + 1, 3, True, True, "Codigo |Color |Cantidad"

If Not consulta.RecordCount = 0 Then
i = 1
Do Until consulta.EOF
 With flehilo
 .Row = i
 .Col = 0
 .Text = i
 
 .Col = 1
 .Text = consulta.Fields("codigo")
 
 .Col = 2
 .Text = consulta.Fields("color")
 
 .Col = 3
 .Text = consulta.Fields("cantidad")
 End With
i = i + 1
consulta.MoveNext
Loop
End If

End Sub
