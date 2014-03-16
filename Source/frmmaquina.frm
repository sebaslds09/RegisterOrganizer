VERSION 5.00
Object = "{4EC837EB-4201-4C6D-A7D7-0C99D69CD9A6}#1.0#0"; "GradientCommand.ocx"
Object = "{13592B48-653C-491D-ACB1-C3140AA12F33}#6.1#0"; "ubGrid.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Begin VB.Form frmmaquina 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Maquina"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   7230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Begin TabDlg.SSTab sstmaquina 
         Height          =   4575
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   7215
         _ExtentX        =   12726
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
         TabPicture(0)   =   "frmmaquina.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "cmdguardar"
         Tab(0).Control(1)=   "cmdcancelar"
         Tab(0).Control(2)=   "cbocodigotipo"
         Tab(0).Control(3)=   "RichTextBox1"
         Tab(0).Control(4)=   "txtnumero"
         Tab(0).Control(5)=   "cbotipo"
         Tab(0).Control(6)=   "txtcodigo"
         Tab(0).Control(7)=   "Label3"
         Tab(0).Control(8)=   "Label2"
         Tab(0).Control(9)=   "Label1"
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "Modificar"
         TabPicture(1)   =   "frmmaquina.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cmdmodificar"
         Tab(1).Control(1)=   "cmdcancelarmod"
         Tab(1).Control(2)=   "cbocodigonuevo"
         Tab(1).Control(3)=   "RichTextBox2"
         Tab(1).Control(4)=   "txtnuevonumero"
         Tab(1).Control(5)=   "cbonuevotipo"
         Tab(1).Control(6)=   "cbomaquina"
         Tab(1).Control(7)=   "Label6"
         Tab(1).Control(8)=   "Label5"
         Tab(1).Control(9)=   "Label4"
         Tab(1).ControlCount=   10
         TabCaption(2)   =   "Eliminar"
         TabPicture(2)   =   "frmmaquina.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmdcancelareli"
         Tab(2).Control(1)=   "cmdeliminar"
         Tab(2).Control(2)=   "RichTextBox3"
         Tab(2).Control(3)=   "cbomaquinaeli"
         Tab(2).Control(4)=   "Label7"
         Tab(2).ControlCount=   5
         TabCaption(3)   =   "Ver"
         TabPicture(3)   =   "frmmaquina.frx":0054
         Tab(3).ControlEnabled=   -1  'True
         Tab(3).Control(0)=   "flemaquina"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).ControlCount=   1
         Begin GradientCommand.GGCommand cmdcancelarmod 
            Height          =   495
            Left            =   -72120
            TabIndex        =   19
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
         Begin GradientCommand.GGCommand cmdmodificar 
            Height          =   495
            Left            =   -74280
            TabIndex        =   18
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
         Begin GradientCommand.GGCommand cmdcancelar 
            Height          =   495
            Left            =   -72480
            TabIndex        =   10
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
            TabIndex        =   9
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
         Begin GradientCommand.GGCommand cmdeliminar 
            Height          =   495
            Left            =   -73920
            TabIndex        =   23
            Top             =   2640
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
            Top             =   2640
            Width           =   1095
            _ExtentX        =   1931
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
         Begin VB.ComboBox cbocodigonuevo 
            Height          =   315
            Left            =   -70800
            TabIndex        =   27
            Top             =   2160
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.ComboBox cbocodigotipo 
            Height          =   315
            Left            =   -71160
            TabIndex        =   26
            Top             =   1680
            Visible         =   0   'False
            Width           =   495
         End
         Begin RichTextLib.RichTextBox RichTextBox3 
            Height          =   3855
            Left            =   -70200
            TabIndex        =   25
            Top             =   480
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   6800
            _Version        =   393217
            BackColor       =   -2147483633
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            TextRTF         =   $"frmmaquina.frx":0070
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
         Begin Proyecto1.SComboBox cbomaquinaeli 
            Height          =   450
            Left            =   -72600
            TabIndex        =   22
            Top             =   1440
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
            Left            =   -70200
            TabIndex        =   20
            Top             =   480
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   6800
            _Version        =   393217
            BackColor       =   -2147483633
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            TextRTF         =   $"frmmaquina.frx":0137
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
         Begin VB.TextBox txtnuevonumero 
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
            Left            =   -72360
            TabIndex        =   17
            Top             =   1560
            Width           =   855
         End
         Begin Proyecto1.SComboBox cbonuevotipo 
            Height          =   450
            Left            =   -72360
            TabIndex        =   16
            Top             =   2040
            Width           =   1455
            _ExtentX        =   2566
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
         Begin Proyecto1.SComboBox cbomaquina 
            Height          =   450
            Left            =   -72360
            TabIndex        =   15
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
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
            Height          =   3975
            Left            =   -70200
            TabIndex        =   11
            Top             =   480
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   7011
            _Version        =   393217
            BackColor       =   -2147483633
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            TextRTF         =   $"frmmaquina.frx":0210
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
         Begin VB.TextBox txtnumero 
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
            TabIndex        =   8
            Top             =   2160
            Width           =   1575
         End
         Begin Proyecto1.SComboBox cbotipo 
            Height          =   450
            Left            =   -72840
            TabIndex        =   7
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
            OfficeAppearance=   2
            ShadowColorText =   6582129
            Style           =   1
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
            ForeColor       =   &H00000000&
            Height          =   450
            Left            =   -72840
            TabIndex        =   6
            Top             =   1200
            Width           =   735
         End
         Begin ubGridControl.ubGrid flemaquina 
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
            Caption         =   "Maquina:"
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
            Left            =   -73920
            TabIndex        =   21
            Top             =   1560
            Width           =   1170
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Nuevo Tipo:"
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
            Left            =   -74520
            TabIndex        =   14
            Top             =   2160
            Width           =   1500
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Nuevo Numero#:"
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
            Left            =   -74520
            TabIndex        =   13
            Top             =   1680
            Width           =   2070
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Maquina:"
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
            Left            =   -74520
            TabIndex        =   12
            Top             =   1080
            Width           =   1170
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Numero #:"
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
            Top             =   2280
            Width           =   1290
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
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
            TabIndex        =   4
            Top             =   1680
            Width           =   645
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
      Bmp:1           =   "frmmaquina.frx":02E3
      Key:1           =   "#mnaarchivosalir"
      Bmp:2           =   "frmmaquina.frx":070B
      Key:2           =   "#mnuayudaayuda"
      Bmp:3           =   "frmmaquina.frx":0B33
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
Attribute VB_Name = "frmmaquina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cargarcodigos()
cargarcodigo "maquina", Me
End Sub
Private Sub cbomaquina_SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
cargartipo
End Sub

Private Sub cbonuevotipo_SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
cbocodigonuevo.ListIndex = cbonuevotipo.ListIndex - 1
End Sub

Private Sub cbotipo_SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
cbocodigotipo.ListIndex = cbotipo.ListIndex - 1
End Sub

Private Sub cmdcancelar_Click()
cbotipo.ListIndex = 0
cbotipo.Text = ""
txtnumero.Text = ""
End Sub

Private Sub cmdcancelareli_Click()
cbomaquinaeli.Text = ""
End Sub

Private Sub cmdcancelarmod_Click()
cargarmaquina
cargartipo
txtnuevonumero.Text = ""
End Sub

Private Sub cmdeliminar_Click()
If cbomaquinaeli.Text = "" Then
 MsgBox ("Elija una maquina"), vbExclamation, "Elejir"
 Else
  If MsgBox("Desea eliminar esta maquina?", vbQuestion + vbYesNo, "Eliminar") = vbYes Then
   sentencia = "delete from maquina where maquina='" & cbomaquinaeli.Text & "'"
   conexion.Execute sentencia
   MsgBox ("La maquina fue eliminada exitosamente"), vbInformation, "Exito"
   cargarmaquina
   cargartipo
   cargarcodigos
   cargarflemaquina
  End If
End If
End Sub
Private Sub cmdguardar_Click()
If cbotipo.Text = "" Then
 MsgBox ("Elija un tipo de maquina"), vbExclamation, "Elejir"
 Else
  If txtnumero.Text = "" Then
   MsgBox ("Escriba un numero"), vbExclamation, "Escribir"
   Else
    If MsgBox("Desea guardar esta maquina?", vbQuestion + vbYesNo, "Guardar") = vbYes Then
    maquina.AddNew
    maquina.Fields("codigo") = txtcodigo.Text
    maquina.Fields("maquina") = "#" & txtnumero.Text
    maquina.Fields("tipo") = cbocodigotipo.Text
    maquina.Update
    MsgBox ("La maquina fue guardada con exito"), vbInformation, "Exito"
    txtnumero.Text = ""
    cargartipo
    cargarcodigos
    cargarflemaquina
    End If
  End If
End If
End Sub
Private Sub cmdmodificar_Click()
If cbomaquina.Text = "" Then
 MsgBox ("Elija una maquina"), vbExclamation, "Elejir"
 Else
 If txtnuevonumero.Text = "" Then
  MsgBox ("Escriba un numero de maquina"), vbExclamation, "Escribir"
   Else
   If cbonuevotipo.Text = "" Then
    MsgBox ("Elija un tipo de maquina"), vbExclamation, "Elejir"
    Else
    If MsgBox("Desea modificar esta maquina?", vbQuestion + vbYesNo, "Moificar") = vbYes Then
    sentencia = "update maquina set maquina='" & "#" & txtnuevonumero.Text & "'" & ", tipo=" & cbocodigonuevo.Text & " where maquina='" & cbomaquina.Text & "'"
    conexion.Execute sentencia
    MsgBox ("La maquina fue modificada exitosamente"), vbInformation, "Exito"
    cargarmaquina
    cargartipo
    cargarflemaquina
    txtnuevonumero.Text = ""
    End If
   End If
 End If
End If
End Sub
Private Sub flemaquina_Click()
flemaquina.SelectionMode = SelectionByRow
End Sub

Private Sub Form_Load()
abrirtablamaquina
abrirtablatipo_maquina
cargarmaquina
cargarcodigos
cargartipo
cargarflemaquina
autoajustar
cargaricono Me
End Sub
Private Sub cargartipo()
cerrarconsulta
sentencia = "select * from tipo_maquina"
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic

cbotipo.Clear
cbocodigotipo.Clear
cbonuevotipo.Clear
cbocodigonuevo.Clear

If Not consulta.RecordCount = 0 Then
Do Until consulta.EOF
 cbotipo.AddItem (consulta.Fields("tipo"))
 cbocodigotipo.AddItem (consulta.Fields("codigo"))
 cbonuevotipo.AddItem (consulta.Fields("tipo"))
 cbocodigonuevo.AddItem (consulta.Fields("codigo"))
 consulta.MoveNext
Loop
consulta.MoveFirst
cbotipo.ListIndex = 1
cbocodigotipo.ListIndex = 0
cbotipo.Text = consulta.Fields("tipo")
cbonuevotipo.ListIndex = 1
cbocodigonuevo.ListIndex = 0
cbonuevotipo.Text = ""
cerrarconsulta
sentencia = "select * from tipo_maquina, maquina where maquina.tipo=tipo_maquina.codigo and maquina='" & cbomaquina.Text & "'"
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic
cbonuevotipo.Text = consulta.Fields("tipo_maquina.tipo")
End If
End Sub

Private Sub mnuarchivosalir_Click()
Unload Me
End Sub

Private Sub sstmaquina_Click(PreviousTab As Integer)
If sstmaquina.Caption = "Nuevo" Then
 cargarcodigos
 cargartipo
End If
If sstmaquina.Caption = "Modificar" Then
 cargartipo
 cargarmaquina
End If
If sstmaquina.Caption = "Eliminar" Then
 cargarmaquina
End If
If sstmaquina.Caption = "Ver" Then
 cargarflemaquina
 autoajustar
End If
End Sub

Private Sub txtnuevonumero_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Then
 If keyascii = 13 Then
  cmdmodificar_Click
  keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub txtnumero_keypress(keyascii As Integer)
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
Private Sub cargarmaquina()
cerrarconsulta
sentencia = "select maquina from maquina"
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic

cbomaquina.Clear
cbomaquinaeli.Clear

If Not consulta.RecordCount = 0 Then
Do Until consulta.EOF
cbomaquina.AddItem (consulta.Fields("maquina"))
cbomaquinaeli.AddItem (consulta.Fields("maquina"))
consulta.MoveNext
Loop
consulta.MoveFirst
cbomaquina.ListIndex = 1
cbomaquinaeli.ListIndex = 1
cbomaquina.Text = consulta.Fields("maquina")
cbomaquinaeli.Text = consulta.Fields("maquina")
End If

End Sub
Private Sub cargarflemaquina()
Dim i As Integer
cerrarconsulta
sentencia = "select * from maquina, tipo_maquina where tipo_maquina.codigo=maquina.tipo"
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic


flemaquina.AutoSetup consulta.RecordCount + 1, 3, True, True, "Codigo |Maquina |Tipo"

If Not consulta.RecordCount = 0 Then
i = 1
Do Until consulta.EOF
 With flemaquina
 .Row = i
 .Col = 0
 .Text = i
 
 .Col = 1
 .Text = consulta.Fields("maquina.codigo")
 
 .Col = 2
 .Text = consulta.Fields("maquina")
 
 .Col = 3
 .Text = consulta.Fields("tipo_maquina.tipo")
 End With
 consulta.MoveNext
 i = i + 1
Loop
End If

End Sub
Private Sub autoajustar()
With flemaquina
.ColWidth(1) = 50
.ColWidth(2) = 60
.ColWidth(3) = 70
End With
End Sub
