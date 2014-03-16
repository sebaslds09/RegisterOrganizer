Attribute VB_Name = "Module2"
Option Explicit
  
Private Sub Main()
conectar
     abrirtablasuiche
  'Carga y visualiza el Splash
  frmSplash.Show
     
  ' Carga en memoria el formulario principal _
    pero no lo muestra
  Load frmsuiche
  Load frmusuario
     
  'Hasta que no se cumpla el tiempo no prosigue
  Do
    DoEvents
  Loop Until frmSplash.Listo
     
     
  ' Visualiza el Form Principal y descarga el Splash
  If suiche.Fields("suiche") = 0 Then
   frmsuiche.Show
   Else
  frmusuario.Show
  End If
  
  Unload frmSplash
  Set frmSplash = Nothing
  
End Sub
Sub Esperar(Tiempo As Single)
Dim TiempoPausa, Inicio, Final, TiempoTotal
   TiempoPausa = Tiempo   ' Asigna hora de inicio.
   Inicio = Timer   ' Establece la hora de inicio.
   Do While Timer < Inicio + TiempoPausa
      DoEvents   ' Cambia a otros procesos.
   Loop
   Final = time   ' Asigna hora de finalización.
   TiempoTotal = Final - Inicio   ' Calcula tiempo total.
End Sub






