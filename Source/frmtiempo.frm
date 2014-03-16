VERSION 5.00
Object = "{4EC837EB-4201-4C6D-A7D7-0C99D69CD9A6}#1.0#0"; "GradientCommand.ocx"
Object = "{13592B48-653C-491D-ACB1-C3140AA12F33}#6.1#0"; "ubGrid.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Begin VB.Form frmtiempo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tiempo Muerto"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   7035
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
      Begin TabDlg.SSTab ssttiempo 
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
         TabPicture(0)   =   "frmtiempo.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(1)=   "Label2"
         Tab(0).Control(2)=   "Label3"
         Tab(0).Control(3)=   "txtcodigo"
         Tab(0).Control(4)=   "cbocausa"
         Tab(0).Control(5)=   "txtmotivo"
         Tab(0).Control(6)=   "RichTextBox1"
         Tab(0).Control(7)=   "cbocodigocausa"
         Tab(0).Control(8)=   "cmdguardar"
         Tab(0).Control(9)=   "cmdcancelar"
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "Modificar"
         TabPicture(1)   =   "frmtiempo.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label4"
         Tab(1).Control(1)=   "Label5"
         Tab(1).Control(2)=   "Label6"
         Tab(1).Control(3)=   "cbomotivo"
         Tab(1).Control(4)=   "cbocausamod"
         Tab(1).Control(5)=   "txtmotivomod"
         Tab(1).Control(6)=   "RichTextBox2"
         Tab(1).Control(7)=   "cbocodigocausamod"
         Tab(1).Control(8)=   "cmdcancelarmod"
         Tab(1).Control(9)=   "cmdmodificar"
         Tab(1).ControlCount=   10
         TabCaption(2)   =   "Eliminar"
         TabPicture(2)   =   "frmtiempo.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label7"
         Tab(2).Control(1)=   "cbomotivoeli"
         Tab(2).Control(2)=   "RichTextBox3"
         Tab(2).Control(3)=   "cmdcancelareli"
         Tab(2).Control(4)=   "cmdeliminar"
         Tab(2).ControlCount=   5
         TabCaption(3)   =   "Ver"
         TabPicture(3)   =   "frmtiempo.frx":0054
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "fletiempo"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).ControlCount=   1
         Begin GradientCommand.GGCommand cmdeliminar 
            Height          =   495
            Left            =   -74160
            TabIndex        =   23
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
         Begin GradientCommand.GGCommand cmdcancelareli 
            Height          =   495
            Left            =   -72120
            TabIndex        =   24
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
            Left            =   -74280
            TabIndex        =   18
            Top             =   3480
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
            Top             =   3480
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
         Begin GradientCommand.GGCommand cmdcancelar 
            Height          =   495
            Left            =   -72600
            TabIndex        =   10
            Top             =   3360
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
            TabIndex        =   9
            Top             =   3360
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
         Begin VB.ComboBox cbocodigocausamod 
            Height          =   315
            Left            =   -70920
            TabIndex        =   27
            Top             =   1800
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.ComboBox cbocodigocausa 
            Height          =   315
            Left            =   -71160
            TabIndex        =   26
            Top             =   1680
            Visible         =   0   'False
            Width           =   495
         End
         Begin RichTextLib.RichTextBox RichTextBox3 
            Height          =   3855
            Left            =   -70440
            TabIndex        =   25
            Top             =   480
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   6800
            _Version        =   393217
            BackColor       =   -2147483633
            ReadOnly        =   -1  'True
            TextRTF         =   $"frmtiempo.frx":0070
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
         Begin Proyecto1.SComboBox cbomotivoeli 
            Height          =   450
            Left            =   -72600
            TabIndex        =   22
            Top             =   1560
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
         Begin RichTextLib.RichTextBox RichTextBox2 
            Height          =   3855
            Left            =   -70440
            TabIndex        =   20
            Top             =   480
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   6800
            _Version        =   393217
            BackColor       =   -2147483633
            ReadOnly        =   -1  'True
            TextRTF         =   $"frmtiempo.frx":016A
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
         Begin VB.TextBox txtmotivomod 
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
            Left            =   -72720
            TabIndex        =   15
            Top             =   2280
            Width           =   1575
         End
         Begin Proyecto1.SComboBox cbocausamod 
            Height          =   450
            Left            =   -72720
            TabIndex        =   14
            Top             =   1800
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
         Begin Proyecto1.SComboBox cbomotivo 
            Height          =   450
            Left            =   -72720
            TabIndex        =   13
            Top             =   1320
            Width           =   2055
            _ExtentX        =   3625
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
            Left            =   -70440
            TabIndex        =   11
            Top             =   480
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   6800
            _Version        =   393217
            BackColor       =   -2147483633
            ReadOnly        =   -1  'True
            TextRTF         =   $"frmtiempo.frx":026C
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
         Begin VB.TextBox txtmotivo 
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
            TabIndex        =   8
            Top             =   2160
            Width           =   1575
         End
         Begin Proyecto1.SComboBox cbocausa 
            Height          =   450
            Left            =   -72840
            TabIndex        =   6
            Top             =   1680
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
            ShadowColorText =   6582129
            Text            =   ""
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
            TabIndex        =   4
            Top             =   1200
            Width           =   975
         End
         Begin ubGridControl.ubGrid fletiempo 
            Height          =   4095
            Left            =   120
            TabIndex        =   2
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
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
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
            Left            =   -74040
            TabIndex        =   21
            Top             =   1680
            Width           =   945
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Nuevo Motivo:"
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
            Left            =   -74760
            TabIndex        =   17
            Top             =   2280
            Width           =   1800
         End
         Begin VB.Label Label5 
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
            Left            =   -74760
            TabIndex        =   16
            Top             =   1800
            Width           =   1710
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
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
            Left            =   -74760
            TabIndex        =   12
            Top             =   1320
            Width           =   945
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
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
            Left            =   -74400
            TabIndex        =   7
            Top             =   2160
            Width           =   945
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
            TabIndex        =   5
            Top             =   1680
            Width           =   855
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
            Height          =   330
            Left            =   -74400
            TabIndex        =   3
            Top             =   1200
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
      Bmp:1           =   "frmtiempo.frx":0336
      Key:1           =   "#mnaarchivosalir"
      Bmp:2           =   "frmtiempo.frx":075E
      Key:2           =   "#mnuayudaayuda"
      Bmp:3           =   "frmtiempo.frx":0B86
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
Attribute VB_Name = "frmtiempo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cargarcodigos()
cargarcodigo "tiempo_muerto", Me
End Sub
Private Sub cbocausa_SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
cbocodigocausa.ListIndex = cbocausa.ListIndex - 1
End Sub
Private Sub cbocausamod_SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
cbocodigocausamod.ListIndex = cbocausamod.ListIndex - 1
End Sub
Private Sub cbomotivo_SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
cargarcausa
End Sub
Private Sub cmdcancelar_Click()
cbotipo.Text = ""
txtnumero.Text = ""
End Sub
Private Sub cmdcancelareli_Click()
cbomaquinaeli.Text = ""
End Sub
Private Sub cmdcancelarmod_Click()
cargarmotivo
cargarcausa
txtmotivomod.Text = ""
End Sub
Private Sub cmdeliminar_Click()
If cbomotivoeli.Text = "" Then
 MsgBox ("Elija un motivo"), vbExclamation, "Elejir"
 Else
  If MsgBox("Desea eliminar este motivo?", vbQuestion + vbYesNo, "Eliminar") = vbYes Then
   sentencia = "delete from tiempo_muerto where motivo='" & cbomotivoeli.Text & "'"
   conexion.Execute sentencia
   MsgBox ("La maquina fue eliminada exitosamente"), vbInformation, "Exito"
   cargarmotivo
   cargarcausa
   cargarcodigos
   cargarfletiempo
  End If
End If
End Sub
Private Sub cmdguardar_Click()
If cbocausa.Text = "" Then
 MsgBox ("Elija una causa"), vbExclamation, "Elejir"
 Else
  If txtmotivo.Text = "" Then
   MsgBox ("Escriba un motivo"), vbExclamation, "Escribir"
   Else
    If MsgBox("Desea guardar este motivo?", vbQuestion + vbYesNo, "Guardar") = vbYes Then
    tiempo_muerto.AddNew
    tiempo_muerto.Fields("codigo") = txtcodigo.Text
    tiempo_muerto.Fields("motivo") = txtmotivo.Text
    tiempo_muerto.Fields("causa") = cbocodigocausa.Text
    tiempo_muerto.Update
    MsgBox ("El motivo fue guardada con exito"), vbInformation, "Exito"
    txtmotivo.Text = ""
    cargarcausa
    cargarcodigos
    cargarfletiempo
    End If
  End If
End If
End Sub
Private Sub cmdmodificar_Click()
If cbomotivo.Text = "" Then
 MsgBox ("Elija un motivo "), vbExclamation, "Elejir"
 Else
 If txtmotivomod.Text = "" Then
  MsgBox ("Escriba un motivo"), vbExclamation, "Escribir"
   Else
   If cbocausamod.Text = "" Then
    MsgBox ("Elija una causa"), vbExclamation, "Elejir"
    Else
    If MsgBox("Desea modificar este motivo?", vbQuestion + vbYesNo, "Moificar") = vbYes Then
    sentencia = "update tiempo_muerto set motivo='" & txtmotivomod.Text & "'" & ", causa=" & cbocodigocausamod.Text & " where motivo='" & cbomotivo.Text & "'"
    conexion.Execute sentencia
    MsgBox ("El motivo fue modificado exitosamente"), vbInformation, "Exito"
    cargarmotivo
    cargarcausa
    cargarfletiempo
    txtmotivomod.Text = ""
    End If
   End If
 End If
End If
End Sub
Private Sub fletipo_Click()
flematipo.SelectionMode = SelectionByRow
End Sub
Private Sub Form_Load()
abrirtablacausa
abrirtablatiempo_muerto
cargarmotivo
cargarcodigos
cargarcausa
cargarfletiempo
autoajustar
cargaricono Me
End Sub
Private Sub cargarcausa()
cerrarconsulta
sentencia = "select * from causa"
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic

cbocausa.Clear
cbocodigocausa.Clear
cbocausamod.Clear
cbocodigocausamod.Clear

If Not consulta.RecordCount = 0 Then
Do Until consulta.EOF
 cbocausa.AddItem (consulta.Fields("causa"))
 cbocodigocausa.AddItem (consulta.Fields("codigo"))
 cbocausamod.AddItem (consulta.Fields("causa"))
 cbocodigocausamod.AddItem (consulta.Fields("codigo"))
 consulta.MoveNext
Loop
consulta.MoveFirst
cbocausa.ListIndex = 1
cbocodigocausa.ListIndex = 0
cbocausa.Text = consulta.Fields("causa")
cbocausamod.ListIndex = 1
cbocodigocausamod.ListIndex = 0
cbocausamod.Text = ""
cerrarconsulta
sentencia = "select * from causa, tiempo_muerto where causa.codigo=tiempo_muerto.causa and motivo='" & cbomotivo.Text & "'"
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic
If Not consulta.RecordCount = 0 Then
cbocausamod.Text = consulta.Fields("causa.causa")
cbocodigocausamod.Text = consulta.Fields("tiempo_muerto.causa")
End If
End If
End Sub

Private Sub mnuarchivosalir_Click()
Unload Me
End Sub

Private Sub ssttiempo_Click(PreviousTab As Integer)
If ssttiempo.Caption = "Nuevo" Then
 cargarcodigos
 cargarcausa
End If
If ssttiempo.Caption = "Modificar" Then
 cargarcausa
 cargarmotivo
End If
If ssttiempo.Caption = "Eliminar" Then
 cargarmotivo
End If
If ssttiempo.Caption = "Ver" Then
 cargarfletiempo
 autoajustar
End If
End Sub
Private Sub txtmotivo_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Or keyascii = 241 Or keyascii = 209 Then
 If keyascii = 13 Then
 cmdguardar_Click
keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo letras y numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub txtmotivomod_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Or keyascii = 241 Or keyascii = 209 Then
 If keyascii = 13 Then
 cmdmodificar_Click
keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo letras y numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub cargarmotivo()
cerrarconsulta
sentencia = "select motivo from tiempo_muerto"
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic

cbomotivo.Clear
cbomotivoeli.Clear

If Not consulta.RecordCount = 0 Then
consulta.MoveFirst

Do Until consulta.EOF
cbomotivo.AddItem (consulta.Fields("motivo"))
cbomotivoeli.AddItem (consulta.Fields("motivo"))
consulta.MoveNext
Loop
consulta.MoveFirst
cbomotivo.ListIndex = 1
cbomotivoeli.ListIndex = 1
cbomotivo.Text = consulta.Fields("motivo")
cbomotivoeli.Text = consulta.Fields("motivo")
End If

End Sub
Private Sub cargarfletiempo()
Dim i As Integer
cerrarconsulta
sentencia = "select * from causa, tiempo_muerto where tiempo_muerto.causa=causa.codigo"
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic

fletiempo.AutoSetup consulta.RecordCount + 1, 3, True, True, "Codigo |Motivo |Causa"

If Not consulta.RecordCount = 0 Then
i = 1
Do Until consulta.EOF
 With fletiempo
 .Row = i
 .Col = 0
 .Text = i
 
 .Col = 1
 .Text = consulta.Fields("tiempo_muerto.codigo")
 
 .Col = 2
 .Text = consulta.Fields("motivo")
 
 .Col = 3
 .Text = consulta.Fields("causa.causa")
 End With
 consulta.MoveNext
 i = i + 1
Loop
End If

End Sub
Private Sub autoajustar()
With fletiempo
.ColWidth(1) = 50
.ColWidth(2) = 100
.ColWidth(3) = 100
End With
End Sub
