VERSION 5.00
Object = "{DA729E34-689F-49EA-A856-B57046630B73}#1.0#0"; "progressbar-xp.ocx"
Begin VB.Form frmSplash 
   BackColor       =   &H00EECB82&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5460
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   10470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Inicio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Inicio.frx":0CCA
   ScaleHeight     =   5460
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00EECB82&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5130
      Left            =   120
      MouseIcon       =   "Inicio.frx":0FD4
      TabIndex        =   0
      Top             =   120
      Width           =   10200
      Begin Proyecto2.XP_ProgressBar xprb1 
         Height          =   255
         Left            =   6720
         TabIndex        =   6
         Top             =   4800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BrushStyle      =   0
         Color           =   12937777
         Orientation     =   1
         Scrolling       =   5
         ShowText        =   -1  'True
      End
      Begin VB.Label lbldir 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4680
         TabIndex        =   5
         Top             =   4800
         Width           =   45
      End
      Begin VB.Label lblWarning 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Todos los derecho recervados"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   4800
         Width           =   2235
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Versión 1.0.3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8565
         TabIndex        =   3
         Top             =   4680
         Width           =   1470
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Autorizado a: Bordados Marion."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   7440
         TabIndex        =   1
         Top             =   240
         Width           =   2490
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Register Organizer"
         BeginProperty Font 
            Name            =   "Edwardian Script ITC"
            Size            =   36
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   5400
         TabIndex        =   4
         Top             =   720
         Width           =   4455
      End
      Begin VB.Image imgLogo 
         Height          =   4425
         Left            =   120
         MouseIcon       =   "Inicio.frx":12DE
         Picture         =   "Inicio.frx":15E8
         Stretch         =   -1  'True
         Top             =   240
         Width           =   9975
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Option Explicit
  
  
'Variable tipo Flag que indica cuando se cumplíó el _
 tiempo para descargar la pantalla de presentación
Public Listo As Boolean
  
'Valor en segundos para la duración de del Splash
Private Const DURACION As Long = 0.1
  
Private Sub Form_Load()
  Dim El_Tiempo As Long, archivo As String
  abrirtablasuiche
     
  Me.Show
  'Almacena el punto de partida para hacer _
   el retardo en segundos
  El_Tiempo = Timer + DURACION
     
  ' ...Hasta que no se cumpla el tiempo realiza el retardo
  Do While Timer < El_Tiempo
     DoEvents
  Loop
     
  ' Pone el Flag en True para descargar este form desde el _
   módulo ( Procedimiento Sub Main)
  Listo = True
  i = 0
archivo = Dir(App.Path & "\*.*")
      Do Until archivo = ""
     
      lbldir.Caption = archivo
      archivo = Dir
      Esperar 0.05
      xprb1.Value = i + 1
      i = xprb1.Value
      Loop
      
      
  archivo = Dir(App.Path & "\ocx\*.*")
        
   
    Do Until archivo = ""
    
     lbldir.Caption = archivo
     archivo = Dir
     Esperar 0.05
     xprb1.Value = i + 1
    i = xprb1.Value
    Loop
    
    archivo = Dir(App.Path & "\base de datos\*.*")
    
   Do Until archivo = ""
     lbldir.Caption = archivo
     archivo = Dir
     Esperar 0.05
     xprb1.Value = i + 1
     i = xprb1.Value
Loop

archivo = Dir(App.Path & "\logos\*.*")
    
  Do Until archivo = ""
     lbldir.Caption = archivo
     archivo = Dir
     Esperar 0.05
     xprb1.Value = i + 1
     i = xprb1.Value
     Loop
For i = i To xprb1.Max
 xprb1.Value = xprb1.Value + 1
Next

End Sub

