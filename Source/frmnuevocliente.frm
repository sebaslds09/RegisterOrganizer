VERSION 5.00
Object = "{4EC837EB-4201-4C6D-A7D7-0C99D69CD9A6}#1.0#0"; "GradientCommand.ocx"
Object = "{13592B48-653C-491D-ACB1-C3140AA12F33}#6.1#0"; "ubGrid.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Begin VB.Form frmnuevocliente 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cliente"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   7020
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
      Begin TabDlg.SSTab sstcliente 
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
         TabPicture(0)   =   "frmnuevocliente.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "cmdguardar"
         Tab(0).Control(1)=   "cmdcancelar"
         Tab(0).Control(2)=   "RichTextBox1"
         Tab(0).Control(3)=   "txttelefono"
         Tab(0).Control(4)=   "txtnombre"
         Tab(0).Control(5)=   "txtcodigo"
         Tab(0).Control(6)=   "Label3"
         Tab(0).Control(7)=   "Label2"
         Tab(0).Control(8)=   "Label1"
         Tab(0).ControlCount=   9
         TabCaption(1)   =   "Modificar"
         TabPicture(1)   =   "frmnuevocliente.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdcancelarmod"
         Tab(1).Control(1)=   "cmdguardarmod"
         Tab(1).Control(2)=   "RichTextBox2"
         Tab(1).Control(3)=   "txtnuevotelefono"
         Tab(1).Control(4)=   "txtnuevonombre"
         Tab(1).Control(5)=   "cbonombre"
         Tab(1).Control(6)=   "Label6"
         Tab(1).Control(7)=   "Label5"
         Tab(1).Control(8)=   "Nombre"
         Tab(1).ControlCount=   9
         TabCaption(2)   =   "Eliminar"
         TabPicture(2)   =   "frmnuevocliente.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmdcancelareli"
         Tab(2).Control(1)=   "cmdeliminar"
         Tab(2).Control(2)=   "RichTextBox3"
         Tab(2).Control(3)=   "cbonombreeli"
         Tab(2).Control(4)=   "Label4"
         Tab(2).ControlCount=   5
         TabCaption(3)   =   "Ver"
         TabPicture(3)   =   "frmnuevocliente.frx":0054
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "flecliente"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).ControlCount=   1
         Begin GradientCommand.GGCommand cmdcancelar 
            Height          =   495
            Left            =   -72120
            TabIndex        =   10
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
         Begin GradientCommand.GGCommand cmdguardar 
            Height          =   495
            Left            =   -74400
            TabIndex        =   9
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
         Begin GradientCommand.GGCommand cmdguardarmod 
            Height          =   495
            Left            =   -74160
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
            ToolTipForeColor=   16711680
            GradientToColor =   15649666
         End
         Begin GradientCommand.GGCommand cmdcancelarmod 
            Height          =   495
            Left            =   -72000
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
            ToolTipForeColor=   16711680
            GradientToColor =   15649666
         End
         Begin GradientCommand.GGCommand cmdeliminar 
            Height          =   495
            Left            =   -74160
            TabIndex        =   23
            Top             =   2400
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
            Left            =   -72000
            TabIndex        =   24
            Top             =   2400
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
            Left            =   -70440
            TabIndex        =   25
            Top             =   480
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   6800
            _Version        =   393217
            BackColor       =   -2147483633
            ReadOnly        =   -1  'True
            TextRTF         =   $"frmnuevocliente.frx":0070
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
            Height          =   495
            Left            =   -72960
            TabIndex        =   22
            Top             =   1080
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   873
            AppearanceCombo =   16
            AutoCompleteWord=   -1  'True
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
            Left            =   -70320
            TabIndex        =   20
            Top             =   480
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   6800
            _Version        =   393217
            BackColor       =   -2147483633
            ReadOnly        =   -1  'True
            TextRTF         =   $"frmnuevocliente.frx":012C
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
         Begin VB.TextBox txtnuevotelefono 
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
            Left            =   -72840
            TabIndex        =   17
            Top             =   2040
            Width           =   2295
         End
         Begin VB.TextBox txtnuevonombre 
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
            Left            =   -72840
            TabIndex        =   15
            Top             =   1560
            Width           =   2295
         End
         Begin Proyecto1.SComboBox cbonombre 
            Height          =   495
            Left            =   -72840
            TabIndex        =   12
            Top             =   840
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   873
            AppearanceCombo =   16
            AutoCompleteWord=   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            GradientColor2  =   15649666
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
            TextRTF         =   $"frmnuevocliente.frx":01F8
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
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   -73080
            TabIndex        =   8
            Top             =   1920
            Width           =   2175
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
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   -73080
            TabIndex        =   6
            Top             =   1440
            Width           =   2175
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
            Left            =   -73080
            TabIndex        =   4
            Top             =   960
            Width           =   1215
         End
         Begin ubGridControl.ubGrid flecliente 
            Height          =   4095
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   6855
            _ExtentX        =   12091
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
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   -74040
            TabIndex        =   21
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   -75000
            TabIndex        =   16
            Top             =   2160
            Width           =   2010
         End
         Begin VB.Label Label5 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   -74880
            TabIndex        =   14
            Top             =   1560
            Width           =   2055
         End
         Begin VB.Label Nombre 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   -74040
            TabIndex        =   13
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   -74400
            TabIndex        =   7
            Top             =   1920
            Width           =   1155
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   -74400
            TabIndex        =   5
            Top             =   1440
            Width           =   1065
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
            Left            =   -74400
            TabIndex        =   3
            Top             =   960
            Width           =   960
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
      Bmp:1           =   "frmnuevocliente.frx":02C2
      Key:1           =   "#mnaarchivosalir"
      Bmp:2           =   "frmnuevocliente.frx":06EA
      Key:2           =   "#mnuayudaayuda"
      Bmp:3           =   "frmnuevocliente.frx":0B12
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
Attribute VB_Name = "frmnuevocliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cargarnombre()
cbonombre.Clear
cbonombreeli.Clear
cerrarconsulta
consultar "select cliente from cliente"

If Not consulta.RecordCount = 0 Then
Do Until consulta.EOF
 cbonombre.AddItem (consulta.Fields("cliente"))
 cbonombreeli.AddItem (consulta.Fields("cliente"))
 consulta.MoveNext
Loop
consulta.MoveFirst
cbonombre.Text = consulta.Fields("cliente")
End If

End Sub

Private Sub cbonombre_SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
cargartxt
End Sub

Private Sub cmdcancelar_Click()
txtnombre.Text = ""
txttelefono.Text = ""
End Sub

Private Sub cmdcancelareli_Click()
cbonombreeli.Text = ""
End Sub

Private Sub cmdcancelarmod_Click()
txtnuevonombre.Text = ""
txtnuevotelefono.Text = ""
End Sub

Private Sub cmdeliminar_Click()
If cbonombreeli.Text = "" Then
MsgBox ("Elija un nombre"), vbExclamation, "Elejir"
 Else
 If MsgBox("Desea eliminar este cliente?", vbQuestion + vbYesNo, "Eliminar") = vbYes Then
  sentencia = "delete from cliente where cliente='" & cbonombreeli.Text & "'"
  conexion.Execute sentencia
  MsgBox ("El cliente fue eliminado exitosamente"), vbInformation, "Exito"
  cargarnombre
  cargarflecliente
 End If
End If
End Sub
Private Sub cmdguardar_Click()
If txtnombre.Text = "" Then
 MsgBox ("Escriba un nombre"), vbExclamation, "Escribir"
 Else
      If txttelefono.Text = "" Then
       MsgBox ("Escriba un telefono"), vbExclamation, "Escribir"
       Else
          If MsgBox("Desea guardar este cliente?", vbQuestion + vbYesNo, "Guardar") = vbYes Then
           cliente.AddNew
           cliente.Fields("codigo") = txtcodigo.Text
           cliente.Fields("cliente") = txtnombre.Text
           cliente.Fields("telefono") = txttelefono.Text
           cliente.Update
           MsgBox ("El cliente fue guardado exitosamente"), vbInformation, "Guardado"
           txtnombre.Text = ""
           txttelefono.Text = ""
           cargarcodigos
           cargarnombre
           cargarflecliente
    End If
  End If
End If
End Sub

Private Sub cmdguardarmod_Click()
If txtnuevonombre.Text = "" Then
 MsgBox ("Escriba un nombre"), vbExclamation, "Escribir"
 Else
      If txtnuevotelefono.Text = "" Then
       MsgBox ("Escriba un telefono"), vbExclamation, "Escribir"
       Else
          If MsgBox("Desea modificar este cliente?", vbQuestion + vbYesNo, "Guardar") = vbYes Then
          sentencia = "update cliente set cliente='" & txtnuevonombre.Text & "'" & " where cliente='" & cbonombre.Text & "'"
          conexion.Execute sentencia
          sentencia = "update cliente set telefono='" & txtnuevotelefono.Text & "'" & " where cliente='" & cbonombre.Text & "'"
          conexion.Execute sentencia
          MsgBox ("El cliente fue modificado exitosamente"), vbInformation, "Exito"
          cbonombre.Text = ""
          txtnuevonombre.Text = ""
          txtnuevotelefono.Text = ""
          cargarnombre
          cargarflecliente
          End If
    End If
End If
End Sub

Private Sub Form_Load()
abrirtablacliente
cargarflecliente
cargarnombre
cargarcodigos
autoajustar
cargaricono Me
End Sub


Private Sub GGCommand1_Click()
dtrcliente.Show vbModal
End Sub

Private Sub mnuarchivosalir_Click()
Unload Me
End Sub

Private Sub sstcliente_Click(PreviousTab As Integer)
If sstcliente.Caption = "Modificar" Then
 cargarnombre
 cargartxt
End If
If sstcliente.Caption = "Eliminar" Then
 cargarnombre
End If
If sstcliente.Caption = "Nuevo" Then
 cargarcodigos
End If
If sstcliente.Caption = "Ver" Then
 cargarflecliente
 autoajustar
End If
End Sub
Private Sub cargarcodigos()
cargarcodigo "cliente", Me
End Sub

Private Sub txtnombre_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Or keyascii = 241 Or keyascii = 209 Then
 If keyascii = 13 Then
 txttelefono.SetFocus
keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo letras y numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub txtnuevonombre_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Or keyascii = 241 Or keyascii = 209 Then
 If keyascii = 13 Then
 txtnuevotelefono.SetFocus
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
  cmdguardar_Click
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub txtnuevotelefono_keypress(keyascii As Integer)
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
Private Sub cargartxt()
cerrarconsulta
sentencia = "select * from cliente where cliente='" & cbonombre.Text & "'"
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic

If Not consulta.RecordCount = 0 Then
txtnuevonombre.Text = consulta.Fields("cliente")
txtnuevotelefono.Text = consulta.Fields("telefono")
End If

End Sub
Private Sub cargarflecliente()
Dim i As Integer
cerrarconsulta
sentencia = "Select * from cliente"
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic

flecliente.AutoSetup consulta.RecordCount + 1, 3, True, True, "Codigo |Nombre |Telefono "

If Not consulta.RecordCount = 0 Then
i = 1
Do Until consulta.EOF
 flecliente.Row = i
 flecliente.Col = 0
 flecliente.Text = i
 
 flecliente.Col = 1
 flecliente.Text = consulta.Fields("codigo")
 
 flecliente.Col = 2
 flecliente.Text = consulta.Fields("cliente")
 
 flecliente.Col = 3
 flecliente.Text = consulta.Fields("telefono")
 
 
 i = i + 1
 consulta.MoveNext
Loop
End If

End Sub
Private Sub autoajustar()
With flecliente
.ColWidth(1) = 50
.ColWidth(2) = 100
.ColWidth(3) = 70
End With
End Sub
