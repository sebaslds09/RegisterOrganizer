VERSION 5.00
Object = "{4EC837EB-4201-4C6D-A7D7-0C99D69CD9A6}#1.0#0"; "GradientCommand.ocx"
Begin VB.Form frmusuario 
   BackColor       =   &H00EECB82&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usuario"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin GradientCommand.GGCommand cmdcancelar 
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "Cancelar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientFromColor=   15649666
      GradientToColor =   -2147483634
   End
   Begin GradientCommand.GGCommand cmdaceptar 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      Caption         =   "Aceptar"
      BackColor       =   15649666
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientFromColor=   15649666
      GradientToColor =   -2147483634
   End
   Begin Proyecto1.SComboBox cbonombre 
      Height          =   300
      Left            =   1080
      TabIndex        =   5
      ToolTipText     =   "Nombre"
      Top             =   360
      Width           =   1575
      _extentx        =   2778
      _extenty        =   529
      appearancecombo =   16
      autocompleteword=   -1  'True
      font            =   "Usuario.frx":0000
      maxlistlength   =   -1
      numberitemstoshow=   -1
      officeappearance=   2
      shadowcolortext =   6582129
      style           =   1
      text            =   ""
   End
   Begin VB.TextBox txtclave 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   2
      ToolTipText     =   "Clave"
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   2160
      Left            =   2760
      Picture         =   "Usuario.frx":0030
      Stretch         =   -1  'True
      ToolTipText     =   "Bordados Marion"
      Top             =   120
      Width           =   2880
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
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
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Clave:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "frmusuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Integer
Private Sub cbonombre_SelectionMade(ByVal SelectedItem As String, ByVal SelectedItemIndex As Long)
txtclave.SetFocus
End Sub
Private Sub cmdaceptar_Click()
Dim sw As Integer
sw = 0
Usuario.MoveFirst
Do Until Usuario.EOF
 If (Usuario.Fields("nombre") = cbonombre.Text And Usuario.Fields("clave") = txtclave.Text And Usuario.Fields("tipo") = "Administrador") Then
  sw = 1
 End If
 If (Usuario.Fields("nombre") = cbonombre.Text And Usuario.Fields("clave") = txtclave.Text And Usuario.Fields("tipo") = "Usuario") Then
 sw = 2
 End If
 Usuario.MoveNext
Loop
 If sw = 1 Then
  If st = 0 Then
   frmmenu.Show
   frmmenu.mnuarchivonuevo.Enabled = True
   frmmenu.mnunuevo.Enabled = True
   frmmenu.mnureportegenerar.Enabled = True
   frmmenu.cmdreporte.Enabled = True
   frmmenu.cmdnuevo.Enabled = True
   
   Unload Me
   Else
    If st = 1 Then
    frmnuevousuario.Show vbModal
    Unload Me
     Else
     If st = 2 Then
      sh = 1
      Me.Hide
       Else
       sh = 2
       Me.Hide
       End If
     End If
  End If
 Else
  If sw = 2 Then
   If st = 0 Then
   frmmenu.Show
   frmmenu.mnuarchivonuevo.Enabled = False
   frmmenu.mnunuevo.Enabled = False
   frmmenu.mnureportegenerar.Enabled = False
   frmmenu.cmdreporte.Enabled = False
   frmmenu.cmdnuevo.Enabled = False
   Me.Hide
   End If
   Else
   MsgBox ("La clave no es correcta"), vbCritical, "Incorrecta"
   txtclave.Text = ""
  End If
 End If
cbonombre.ListIndex = -1
txtclave.Text = ""
sw = 0
 
End Sub

Private Sub cmdcancelar_Click()
Unload Me
End Sub


Private Sub Form_Load()
conectar
abrirtablausuario
Do Until Usuario.EOF
 cbonombre.AddItem (Usuario.Fields("nombre"))
 Usuario.MoveNext
Loop
If frmmenu.Visible = True Then
 st = 1
Else
 st = 0
End If
cargaricono Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmSplash
Unload Me
End Sub
Private Sub txtclave_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Or keyascii = 241 Or keyascii = 209 Then
 If keyascii = 13 Then
  cmdaceptar_Click
  keyascii = 0
 End If
Else

 MsgBox ("Ingrese solo letras y numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub

