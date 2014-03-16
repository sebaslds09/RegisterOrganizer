VERSION 5.00
Object = "{4EC837EB-4201-4C6D-A7D7-0C99D69CD9A6}#1.0#0"; "GradientCommand.ocx"
Object = "{13592B48-653C-491D-ACB1-C3140AA12F33}#6.1#0"; "ubGrid.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Begin VB.Form frmtipomaquina 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipo Maquina"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
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
      Begin TabDlg.SSTab ssttipo 
         Height          =   4575
         Left            =   0
         TabIndex        =   1
         Top             =   480
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   8070
         _Version        =   393216
         Style           =   1
         Tabs            =   4
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
         TabPicture(0)   =   "frmtipomaquina.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label3"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "txtcodigo"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txttipo"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "txtcabezotes"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "RichTextBox1"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "cmdguardar"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "cmdcancelar"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).ControlCount=   9
         TabCaption(1)   =   "Modificar"
         TabPicture(1)   =   "frmtipomaquina.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label4"
         Tab(1).Control(1)=   "Label5"
         Tab(1).Control(2)=   "Label6"
         Tab(1).Control(3)=   "txtnuevotipo"
         Tab(1).Control(4)=   "txtnuevacantidad"
         Tab(1).Control(5)=   "cbotipo"
         Tab(1).Control(6)=   "RichTextBox2"
         Tab(1).Control(7)=   "cmdmodificar"
         Tab(1).Control(8)=   "cmdcancelarmod"
         Tab(1).ControlCount=   9
         TabCaption(2)   =   "Eliminar"
         TabPicture(2)   =   "frmtipomaquina.frx":0038
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label7"
         Tab(2).Control(1)=   "cbotipoeli"
         Tab(2).Control(2)=   "RichTextBox3"
         Tab(2).Control(3)=   "cmdcancelareli"
         Tab(2).Control(4)=   "cmdeliminar"
         Tab(2).ControlCount=   5
         TabCaption(3)   =   "Ver"
         TabPicture(3)   =   "frmtipomaquina.frx":0054
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "fletipo"
         Tab(3).ControlCount=   1
         Begin GradientCommand.GGCommand cmdcancelar 
            Height          =   495
            Left            =   2640
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
            Left            =   480
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
         Begin GradientCommand.GGCommand cmdmodificar 
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
         Begin GradientCommand.GGCommand cmdeliminar 
            Height          =   495
            Left            =   -74160
            TabIndex        =   23
            Top             =   2760
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
            Top             =   2760
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
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   6800
            _Version        =   393217
            BackColor       =   -2147483633
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            TextRTF         =   $"frmtipomaquina.frx":0070
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
         Begin Proyecto1.SComboBox cbotipoeli 
            Height          =   450
            Left            =   -72840
            TabIndex        =   22
            Top             =   1560
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
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            TextRTF         =   $"frmtipomaquina.frx":0132
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
         Begin Proyecto1.SComboBox cbotipo 
            Height          =   450
            Left            =   -72840
            TabIndex        =   17
            Top             =   1200
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
            Height          =   450
            Left            =   -72840
            TabIndex        =   16
            Top             =   2160
            Width           =   1335
         End
         Begin VB.TextBox txtnuevotipo 
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
            TabIndex        =   15
            Top             =   1680
            Width           =   1695
         End
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   3855
            Left            =   4560
            TabIndex        =   11
            Top             =   480
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   6800
            _Version        =   393217
            BackColor       =   -2147483633
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            TextRTF         =   $"frmtipomaquina.frx":020B
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
         Begin VB.TextBox txtcabezotes 
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
            Left            =   2040
            TabIndex        =   8
            Top             =   2160
            Width           =   1215
         End
         Begin VB.TextBox txttipo 
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
            Left            =   2040
            TabIndex        =   7
            Top             =   1680
            Width           =   1815
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
            Left            =   2040
            TabIndex        =   6
            Top             =   1200
            Width           =   975
         End
         Begin ubGridControl.ubGrid fletipo 
            Height          =   4095
            Left            =   -74880
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
         Begin VB.Label Label7 
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
            Height          =   450
            Left            =   -73680
            TabIndex        =   21
            Top             =   1680
            Width           =   645
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Cabezotes:"
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
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Label5 
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
            TabIndex        =   13
            Top             =   1680
            Width           =   1500
         End
         Begin VB.Label Label4 
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
            Height          =   330
            Left            =   -74520
            TabIndex        =   12
            Top             =   1200
            Width           =   645
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cabezotes:"
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
            Top             =   2280
            Width           =   1335
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
            Height          =   330
            Left            =   480
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
            Height          =   330
            Left            =   480
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
      Bmp:1           =   "frmtipomaquina.frx":02E2
      Key:1           =   "#mnaarchivosalir"
      Bmp:2           =   "frmtipomaquina.frx":070A
      Key:2           =   "#mnuayudaayuda"
      Bmp:3           =   "frmtipomaquina.frx":0B32
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
Attribute VB_Name = "frmtipomaquina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cargarcodigos()
cargarcodigo "tipo_maquina", Me
End Sub
Private Sub cbotipo_SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
cargartxt
End Sub
Private Sub cmdcancelar_Click()
cargarcodigos
txttipo.Text = ""
txtcabezotes.Text = ""
End Sub


Private Sub cmdcancelarmod_Click()
txtnuevotipo.Text = ""
txtnuevacantidad.Text = ""
cargartipo
End Sub

Private Sub cmdeliminar_Click()
If cbotipoeli.Text = "" Then
MsgBox ("Elija un tipo"), vbExclamation, "Elejir"
 Else
 If MsgBox("Desea eliminar este tipo de maquina?", vbQuestion + vbYesNo, "Eliminar") = vbYes Then
  sentencia = "delete from tipo_maquina where tipo='" & cbotipoeli.Text & "'"
  conexion.Execute sentencia
  MsgBox ("El tipo de maquina fue eliminado exitosamente"), vbInformation, "Exito"
  cargartipo
  cargarfletipo
  cmdcancelarmod_Click
 End If
End If
End Sub
Private Sub cmdmodificar_Click()
If txtnuevotipo.Text = "" Then
 MsgBox ("Escriba un tipo"), vbExclamation, "Escribir"
 Else
      If txtnuevacantidad.Text = "" Then
       MsgBox ("Escriba una cantidad"), vbExclamation, "Escribir"
       Else
          If MsgBox("Desea modificar este tipo de maquina?", vbQuestion + vbYesNo, "Guardar") = vbYes Then
          sentencia = "update tipo_maquina set tipo='" & txtnuevotipo.Text & "'" & " where tipo='" & cbotipo.Text & "'"
          conexion.Execute sentencia
          sentencia = "update tipo_maquina set cabezotes='" & txtnuevacantidad.Text & "'" & " where tipo='" & cbotipo.Text & "'"
          conexion.Execute sentencia
          MsgBox ("El tipo de maquina fue modificado exitosamente"), vbInformation, "Exito"
          cbotipo.Text = ""
          txtnuevotipo.Text = ""
          txtnuevacantidad.Text = ""
          cargartipo
          cargarfletipo
          End If
    End If
End If
End Sub
Private Sub fletipo_Click()
fletipo.SelectionMode = SelectionByRow
End Sub

Private Sub mnuarchivosalir_Click()
Unload Me
End Sub

Private Sub ssttipo_Click(PreviousTab As Integer)
If ssttipo.Caption = "Nuevo" Then
 cargarcodigos
End If
If ssttipo.Caption = "Modificar" Then
 cargartipo
 cargartxt
End If
If ssttipo.Caption = "Eliminar" Then
 cargartipo
End If
If ssttipo.Caption = "Ver" Then
 cargarfletipo
 autoajustar
End If
End Sub
Private Sub txttipo_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Or keyascii = 241 Or keyascii = 209 Then
 If keyascii = 13 Then
 txtcabezotes.SetFocus
keyascii = 0
End If
Else
 MsgBox ("Ingrese solo letras y numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub txtcabezotes_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Then
 If keyascii = 13 Then
  cmdguardar_Click
  keyascii = 0
 End If
Else
 If keyascii = 180 Then
  keyascii = 0
  keyascii = 8
End If
 MsgBox ("Ingrese solo numeros"), vbCritical, "Error"
 keyascii = 0
 keyascii = 8
End If
End Sub
Private Sub Form_Load()
abrirtablatipo_maquina
cargarcodigos
cargartipo
cargarfletipo
cargartxt
autoajustar
cargaricono Me
End Sub
Private Sub cmdguardar_Click()
If txttipo.Text = "" Then
 MsgBox ("Escriba un tipo"), vbExclamation, "Escribir"
 Else
      If txtcabezotes.Text = "" Then
       MsgBox ("Escriba una cantidad"), vbExclamation, "Escribir"
       Else
          If MsgBox("Desea guardar este tipo de maquina?", vbQuestion + vbYesNo, "Guardar") = vbYes Then
           tipo_maquina.AddNew
           tipo_maquina.Fields("codigo") = txtcodigo.Text
           tipo_maquina.Fields("tipo") = txttipo.Text
           tipo_maquina.Fields("cabezotes") = txtcabezotes.Text
           tipo_maquina.Update
           MsgBox ("El tipo de maquina fue guardado exitosamente"), vbInformation, "Guardado"
           txttipo.Text = ""
           txtcabezotes.Text = ""
           cargarcodigos
           cargartipo
           cargarfletipo
    End If
  End If
End If
End Sub
Private Sub cargartipo()
cerrarconsulta
sentencia = "select tipo from tipo_maquina"
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic

cbotipo.Clear
cbotipoeli.Clear

If Not consulta.RecordCount = 0 Then
Do Until consulta.EOF
 cbotipo.AddItem (consulta.Fields("tipo"))
 cbotipoeli.AddItem (consulta.Fields("tipo"))
 consulta.MoveNext
Loop
consulta.MoveFirst
cbotipo.ListIndex = 1
cbotipoeli.ListIndex = 0
cbotipo.Text = consulta.Fields("tipo")
cbotipoeli.Text = consulta.Fields("tipo")
End If

End Sub
Private Sub cargartxt()
cerrarconsulta
sentencia = "select * from tipo_maquina where tipo='" & cbotipo.Text & "'"
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic

If Not consulta.RecordCount = 0 Then
txtnuevotipo.Text = consulta.Fields("tipo")
txtnuevacantidad.Text = consulta.Fields("cabezotes")
End If

End Sub
Private Sub txtnuevotipo_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Or keyascii = 241 Or keyascii = 209 Then
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
Private Sub cargarfletipo()
Dim i As Integer
cerrarconsulta
sentencia = "Select * from tipo_maquina"
consulta.Source = sentencia
consulta.Open , conexion, adOpenStatic

fletipo.AutoSetup consulta.RecordCount + 1, 3, True, True, "Codigo |Tipo |Cabezotes "

If Not consulta.RecordCount = 0 Then
i = 1
Do Until consulta.EOF
 fletipo.Row = i
 fletipo.Col = 0
 fletipo.Text = i
 
 fletipo.Col = 1
 fletipo.Text = consulta.Fields("codigo")
 
 fletipo.Col = 2
 fletipo.Text = consulta.Fields("tipo")
 
 fletipo.Col = 3
 fletipo.Text = consulta.Fields("cabezotes")
 
 
 i = i + 1
 consulta.MoveNext
Loop
End If
End Sub
Private Sub autoajustar()
With fletipo
.ColWidth(1) = 50
.ColWidth(2) = 150
.ColWidth(3) = 80
End With
End Sub

