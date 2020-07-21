Attribute VB_Name = "ModPrincipal"
'***************************************************************************
'* Open Source
'* System Application Software - Funcines virtuales de Búsqueda
'* Módulo ModPrincipal de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************
Private escriptar As New clsEscritar

Sub Main()
 ModWinXP.InitCommonControlsVB
 Abrir
 If y_seguridad.y_inicio = "" Then
 MDIPrincipal.Show
 Else
 frmIniciarSecion.Show
 End If
End Sub

Public Sub Abrir()
modoEscrpto ' Escripto Bytes
On Error GoTo no_se
Open "datosx.sys" For Input As 1
 Do While Not EOF(1)
      Line Input #1, seguridad.x_inicio
       y_seguridad.y_inicio = escriptar.funcion_desescriptar(seguridad.x_inicio)
     '%%%%%%%%%%'
     Line Input #1, seguridad.x_modific
       y_seguridad.y_modific = escriptar.funcion_desescriptar(seguridad.x_modific)
     '%%%%%%%%%%'
     Line Input #1, seguridad.x_creacion
       y_seguridad.y_creacion = escriptar.funcion_desescriptar(seguridad.x_creacion)
     '%%%%%%%%%%'
     Line Input #1, seguridad.x_busqueda
       y_seguridad.y_busqueda = escriptar.funcion_desescriptar(seguridad.x_busqueda)
     '%%%%%%%%%%'
     Line Input #1, seguridad.x_poderver
       y_seguridad.y_poderver = escriptar.funcion_desescriptar(seguridad.x_poderver)
     '%%%%%%%%%%'
     Line Input #1, seguridad.x_iniciodel
       y_seguridad.y_iniciodel = escriptar.funcion_desescriptar(seguridad.x_iniciodel)
     '%%%%%%%%%%'
     Line Input #1, seguridad.x_elimnarTodo
      y_seguridad.y_elimnarTodo = escriptar.funcion_desescriptar(seguridad.x_elimnarTodo)
     '%%%%%%%%%%'
      Line Input #1, seguridad.x_eliminarSeleccionado
      y_seguridad.y_eliminarSeleccionado = escriptar.funcion_desescriptar(seguridad.x_eliminarSeleccionado)
     '%%%%%%%%%%'
      Line Input #1, seguridad.x_poderExportar
      y_seguridad.y_poderExportar = escriptar.funcion_desescriptar(seguridad.x_poderExportar)
     '%%%%%%%%%%'
      Line Input #1, seguridad.x_poderImprimir
      y_seguridad.y_poderImprimir = escriptar.funcion_desescriptar(seguridad.x_poderImprimir)
     '%%%%%%%%%%'
      Line Input #1, seguridad.x_crearCopiaSeguridad
      y_seguridad.y_crearCopiaSeguridad = escriptar.funcion_desescriptar(seguridad.x_crearCopiaSeguridad)
     '%%%%%%%%%%'
     Loop
     Close #1
no_se:
End Sub

Public Sub Guardar()
modoEscrpto ' Escripto Bytes
Open "datosx.sys" For Output As 1
 Print #1, escriptar.funcion_escriptar(y_seguridad.y_inicio)
 Print #1, escriptar.funcion_escriptar(y_seguridad.y_modific)
 Print #1, escriptar.funcion_escriptar(y_seguridad.y_creacion)
 Print #1, escriptar.funcion_escriptar(y_seguridad.y_busqueda)
 Print #1, escriptar.funcion_escriptar(y_seguridad.y_poderver)
 Print #1, escriptar.funcion_escriptar(y_seguridad.y_iniciodel)
 Print #1, escriptar.funcion_escriptar(y_seguridad.y_elimnarTodo)
 Print #1, escriptar.funcion_escriptar(y_seguridad.y_eliminarSeleccionado)
 Print #1, escriptar.funcion_escriptar(y_seguridad.y_poderExportar)
 Print #1, escriptar.funcion_escriptar(y_seguridad.y_poderImprimir)
 Print #1, escriptar.funcion_escriptar(y_seguridad.y_crearCopiaSeguridad)
Close #1
End Sub

Public Sub modoEscrpto()
 escriptar.variable_desescriptar_bytes = Mod_Funciones_conByts.escript_121
 escriptar.variable_escriptar_bytes = Mod_Funciones_conByts.escript_121
End Sub


