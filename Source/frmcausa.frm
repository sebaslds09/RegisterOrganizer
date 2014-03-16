VERSION 5.00
Object = "{4EC837EB-4201-4C6D-A7D7-0C99D69CD9A6}#1.0#0"; "GradientCommand.ocx"
Object = "{13592B48-653C-491D-ACB1-C3140AA12F33}#6.1#0"; "ubGrid.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Begin VB.Form frmcausa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Causa"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin HookMenu.XpMenu XpMenu1 
      Left            =   6720
      Top             =   0
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
      ShortCutSelectColor=   12582912
      ArrowNormalColor=   15649666
      ArrowSelectColor=   16744576
      ShadowColor     =   0
      Bmp:1           =   "frmcausa.frx":0000
      Key:1           =   "#mnuarchivosalir"
      Bmp:2           =   "frmcausa.frx":0428
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
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   9763
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
      Begin TabDlg.SSTab sstcausa 
         Height          =   4575
         Left            =   0
         TabIndex        =   1
         Top             =   360
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
         TabPicture(0)   =   "frmcausa.frx":0850
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "cmdcancelar"
         Tab(0).Control(1)=   "cmdguardar"
         Tab(0).Control(2)=   "RichTextBox1"
         Tab(0).Control(3)=   "txtcausa"
         Tab(0).Control(4)=   "txtcodigo"
         Tab(0).Control(5)=   "Label2"
         Tab(0).Control(6)=   "Label1"
         Tab(0).ControlCount=   7
         TabCaption(1)   =   "Modificar"
         TabPicture(1)   =   "frmcausa.frx":086C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdcancelarmod"
         Tab(1).Control(1)=   "cmdmodificar"
         Tab(1).Control(2)=   "RichTextBox2"
         Tab(1).Control(3)=   "cbocausa"
         Tab(1).Control(4)=   "txtnuevacausa"
         Tab(1).Control(5)=   "Label4"
         Tab(1).Control(6)=   "Label3"
         Tab(1).ControlCount=   7
         TabCaption(2)   =   "Eliminar"
         TabPicture(2)   =   "frmcausa.frx":0888
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label5"
         Tab(2).Control(1)=   "cbocausaeli"
         Tab(2).Control(2)=   "RichTextBox3"
         Tab(2).Control(3)=   "cmdeliminar"
         Tab(2).Control(4)=   "cmdcancelareli"
         Tab(2).ControlCount=   5
         TabCaption(3)   =   "Ver"
         TabPicture(3)   =   "frmcausa.frx":08A4
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "flecausa"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).ControlCount=   1
         Begin GradientCommand.GGCommand cmdguardar 
            Height          =   495
            Left            =   -74520
            TabIndex        =   7
            Top             =   2880
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
            Left            =   -72360
            TabIndex        =   8
            Top             =   2880
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
         Begin GradientCommand.GGCommand cmdmodificar 
            Height          =   495
            Left            =   -74160
            TabIndex        =   14
            Top             =   3000
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
            Left            =   -72480
            TabIndex        =   15
            Top             =   3000
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
         Begin GradientCommand.GGCommand cmdcancelareli 
            Height          =   495
            Left            =   -72360
            TabIndex        =   20
            Top             =   2880
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
            TabIndex        =   19
            Top             =   2880
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
         Begin RichTextLib.RichTextBox RichTextBox3 
            Height          =   3855
            Left            =   -70320
            TabIndex        =   21
            Top             =   480
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   6800
            _Version        =   393217
            BackColor       =   -2147483633
            ReadOnly        =   -1  'True
            TextRTF         =   $"frmcausa.frx":08C0
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
         Begin Proyecto1.SComboBox cbocausaeli 
            Height          =   450
            Left            =   -72840
            TabIndex        =   18
            Top             =   1560
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   794
            AppearanceCombo =   16
            BackColor       =   16120314
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
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
            TabIndex        =   16
            Top             =   480
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   6800
            _Version        =   393217
            BackColor       =   -2147483633
            ReadOnly        =   -1  'True
            TextRTF         =   $"frmcausa.frx":098F
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
         Begin Proyecto1.SComboBox cbocausa 
            Height          =   450
            Left            =   -73080
            TabIndex        =   13
            Top             =   1320
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
            ShadowColorText =   6582129
            Text            =   ""
         End
         Begin VB.TextBox txtnuevacausa 
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
            Left            =   -73080
            TabIndex        =   12
            Top             =   1800
            Width           =   1575
         End
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   3855
            Left            =   -70320
            TabIndex        =   9
            Top             =   480
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   6800
            _Version        =   393217
            BackColor       =   -2147483633
            ReadOnly        =   -1  'True
            TextRTF         =   $"frmcausa.frx":0A3F
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
         Begin VB.TextBox txtcausa 
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
            Left            =   -72960
            TabIndex        =   6
            Top             =   1800
            Width           =   1575
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
            Left            =   -72960
            TabIndex        =   5
            Top             =   1320
            Width           =   735
         End
         Begin ubGridControl.ubGrid flecausa 
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
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
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
            Left            =   -74160
            TabIndex        =   17
            Top             =   1680
            Width           =   855
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nueva Causa:"
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
            Left            =   -74880
            TabIndex        =   11
            Top             =   1920
            Width           =   1710
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
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
            Left            =   -74880
            TabIndex        =   10
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            Left            =   -74400
            TabIndex        =   4
            Top             =   1920
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Codigio:"
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
            Left            =   -74400
            TabIndex        =   3
            Top             =   1320
            Width           =   1035
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
Attribute VB_Name = "frmcausa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cargarcodigos()
cargarcodigo "causa", Me
End Sub
Private Sub cargarflecausa()
cerrarconsulta
sentencia = "select * from causa"
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic

flecausa.AutoSetup consulta.RecordCount + 1, 2, True, True, "Codigo |Causa"

If Not consulta.RecordCount = 0 Then
i = 1
Do Until consulta.EOF
 With flecausa
 .Row = i
 .Col = 0
 .Text = i
 
 .Col = 1
 .Text = consulta.Fields("codigo")
 
 .Col = 2
 .Text = consulta.Fields("causa")
  End With
i = i + 1
consulta.MoveNext
Loop
End If

End Sub
Private Sub cargartxt()
cerrarconsulta
sentencia = "select * from causa where causa='" & cbocausa.Text & "'"
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic

If Not consulta.RecordCount = 0 Then
txtnuevacausa.Text = consulta.Fields("causa")
End If

End Sub
Private Sub cargarcausa()
cerrarconsulta
sentencia = "select causa from causa"
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic

cbocausa.Clear
cbocausaeli.Clear

If Not consulta.RecordCount = 0 Then
Do Until consulta.EOF
 cbocausa.AddItem (consulta.Fields("causa"))
 cbocausaeli.AddItem (consulta.Fields("causa"))
 consulta.MoveNext
Loop
consulta.MoveFirst
cbocausa.Text = consulta.Fields("causa")
End If

End Sub

Private Sub cbocausa_SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
cargartxt
End Sub



Private Sub cmdcancelar_Click()
txtcausa.Text = ""
End Sub

Private Sub cmdcancelareli_Click()
cbocausaeli.Text = ""
End Sub

Private Sub cmdcancelarmod_Click()
txtnuevacausa.Text = ""
End Sub

Private Sub flecausa_Click()
flecausa.SelectionMode = SelectionByRow
End Sub

Private Sub mnuarchivosalir_Click()
Unload Me
End Sub

Private Sub txtcausa_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Or keyascii = 164 Or keyascii = 165 Then
 If keyascii = 13 Then
  cmdguardar_Click
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo letras y numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub txtnuevacausa_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Or keyascii = 164 Or keyascii = 165 Then
 If keyascii = 13 Then
  cmdmodificar_Click
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo letras y numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub sstcausa_Click(PreviousTab As Integer)
If sstcausa.Caption = "Modificar" Then
 cargarcausa
 cargartxt
End If
If sstcausa.Caption = "Nuevo" Then
 cargarcodigos
End If
If sstcausa.Caption = "Eliminar" Then
 cargarcausa
End If
If sstcausa.Caption = "Ver" Then
 cargarflecausa
 autoajustar
End If
End Sub
Private Sub Form_Load()
abrirtablacausa
cargarcodigos
cargarcausa
cargartxt
cargarflecausa
autoajustar
cargaricono Me
End Sub
Private Sub cmdmodificar_Click()
If txtnuevacausa.Text = "" Then
 MsgBox ("Escriba una causa"), vbExclamation, "Escribir"
  Else
     If MsgBox("Desea modificar esta causa?", vbQuestion + vbYesNo, "Modificar") = vbYes Then
     sentencia = "update causa set causa='" & txtnuevacausa.Text & "'" & " where causa='" & cbocausa.Text & "'"
     conexion.Execute sentencia
     MsgBox ("La causa fue modificado exitosamente"), vbInformation, "Exito"
     txtnuevacausa.Text = ""
     cbocausa.Text = ""
     cargarcausa
     cargartxt
     cargarflecausa
    End If
End If
End Sub
Private Sub cmdguardar_Click()
If txtcausa.Text = "" Then
 MsgBox ("Escriba una causa"), vbExclamation, "Escribir"
    Else
     If MsgBox("Desea guardar esta causa", vbQuestion + vbYesNo, "Guardar") = vbYes Then
      causa.AddNew
      causa.Fields("codigo") = txtcodigo.Text
      causa.Fields("causa") = txtcausa.Text
      causa.Update
      MsgBox ("La causa fue guardada con exito"), vbInformation, "Exito"
      txtcausa.Text = ""
      cargarcausa
      cargarflecausa
      cargarcodigos
    End If
End If
End Sub
Private Sub cmdeliminar_Click()
If cbocausaeli.Text = "" Then
 MsgBox ("Elija una causa"), vbExclamation, "Elejir"
 Else
 If MsgBox("Desea eliminar la causa?", vbQuestion + vbYesNo, "Eliminar") = vbYes Then
  sentencia = "delete from causa where causa='" & cbocausaeli.Text & "'"
  conexion.Execute sentencia
  MsgBox ("La causa fue eliminada exitosamente"), vbInformation, "Exito"
  cargarcausa
  cargarflecausa
  cmdcancelarmod_Click
 End If
End If
End Sub
Private Sub autoajustar()
With flecausa
.ColWidth(1) = 50
.ColWidth(2) = 100
End With
End Sub
