VERSION 5.00
Object = "{4EC837EB-4201-4C6D-A7D7-0C99D69CD9A6}#1.0#0"; "GradientCommand.ocx"
Object = "{ADBBAED6-B16F-40DC-80DF-B44910CBA76C}#1.0#0"; "Frame-Xp.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmsuiche 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usuario"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XPFrame.FrameXp FrameXp1 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   13361
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
      Begin GradientCommand.GGCommand cmdsalir 
         Height          =   615
         Left            =   360
         TabIndex        =   7
         Top             =   3240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
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
         GradientToColor =   15649666
      End
      Begin GradientCommand.GGCommand cmdguardar 
         Height          =   615
         Left            =   4080
         TabIndex        =   8
         Top             =   3240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1085
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
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   3495
         Left            =   240
         TabIndex        =   9
         Top             =   3960
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   6165
         _Version        =   393217
         BackColor       =   -2147483633
         ReadOnly        =   -1  'True
         TextRTF         =   $"frmsuiche.frx":0000
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
      Begin VB.TextBox txtclavere 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2880
         TabIndex        =   6
         Top             =   2160
         Width           =   2775
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
         Height          =   495
         Left            =   2880
         TabIndex        =   2
         Top             =   1440
         Width           =   2775
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
         Height          =   495
         Left            =   2880
         TabIndex        =   1
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00EECB82&
         Caption         =   "Repetir Contraseña:"
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
         Left            =   240
         TabIndex        =   5
         Top             =   2160
         Width           =   2490
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00EECB82&
         Caption         =   "Contraseña:"
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
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00EECB82&
         Caption         =   "Nombre de Usuario:"
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
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   2460
      End
   End
End
Attribute VB_Name = "frmsuiche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdguardar_Click()
If txtnombre.Text = "" Then
 MsgBox ("Escriba un nombre"), vbExclamation, "Escribir"
 Else
 If txtclave.Text = "" Then
  MsgBox ("Escriba un nombre"), vbExclamation, "Escribir"
  Else
  If txtclavere.Text = "" Then
   MsgBox ("Escriba un nombre"), vbExclamation, "Escribir"
   Else
   If txtclave.Text = txtclavere.Text Then
   Usuario.AddNew
   Usuario.Fields("nombre") = txtnombre.Text
   Usuario.Fields("clave") = txtclave.Text
   Usuario.Fields("tipo") = "Administrador"
   Usuario.Update
   consultarotro "update suiche set suiche=" & 1
   Unload frmusuario
   frmusuario.Show
   Unload Me
   Else
   MsgBox ("Las claves no son iguales"), vbCritical, "Error"
   txtclave.Text = ""
   txtclavere.Text = ""

   End If
  End If
 End If
End If
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
abrirtablasuiche
abrirtablausuario
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
Unload frmusuario
Unload frmmenu
End Sub
Private Sub txtnombre_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Or keyascii = 241 Or keyascii = 209 Then
 If keyascii = 13 Then
 txtclave.SetFocus
keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo letras y numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub txtclave_keypress(keyascii As Integer)
If (keyascii >= 48 And keyascii <= 57) Or (keyascii >= 65 And keyascii <= 90) Or (keyascii >= 97 And keyascii <= 122) Or keyascii = 32 Or keyascii = 8 Or keyascii = 13 Or keyascii = 241 Or keyascii = 209 Then
 If keyascii = 13 Then
 txtclavere.SetFocus
keyascii = 0
 End If
Else
 MsgBox ("Ingrese solo letras y numeros"), vbCritical, "Error"
 keyascii = 0
End If
End Sub
Private Sub txtclavere_keypress(keyascii As Integer)
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
