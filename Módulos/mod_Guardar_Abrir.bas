Attribute VB_Name = "mod_Guardar_Abrir"
'***************************************************************************
'* Open Source
'* System Application Software
'* Módulo mod_Guardar_Abrir de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************
Public escriptar As New clsEscritar

Public Sub Abrir()
 modoEscrpto ' Escripto Bytes
 On Error GoTo no_se
 Open "bd.opxl" For Input As 1
 Do While Not EOF(1)
 With agenda
 Line Input #1, almacen.v_nombre
 agenda.nombre.Add escriptar.funcion_desescriptar(almacen.v_nombre)
 Line Input #1, almacen.v_nombrex
 agenda.nombrex.Add escriptar.funcion_desescriptar(almacen.v_nombrex)
 Line Input #1, almacen.v_apellidop
 agenda.apellidop.Add escriptar.funcion_desescriptar(almacen.v_apellidop)
 Line Input #1, almacen.v_apellidom
 agenda.apellidom.Add escriptar.funcion_desescriptar(almacen.v_apellidom)
 Line Input #1, almacen.v_telefono0
 agenda.telefono0.Add escriptar.funcion_desescriptar(almacen.v_telefono0)
 Line Input #1, almacen.v_telefono1
 agenda.telefono1.Add escriptar.funcion_desescriptar(almacen.v_telefono1)
 Line Input #1, almacen.v_hora
 agenda.hora.Add escriptar.funcion_desescriptar(almacen.v_hora)
 Line Input #1, almacen.v_celular0
 agenda.celular0.Add escriptar.funcion_desescriptar(almacen.v_celular0)
 Line Input #1, almacen.v_celular1
 agenda.celular1.Add escriptar.funcion_desescriptar(almacen.v_celular1)
 Line Input #1, almacen.v_ci
 agenda.ci.Add escriptar.funcion_desescriptar(almacen.v_ci)
 Line Input #1, almacen.v_direccion0
 agenda.direccion0.Add escriptar.funcion_desescriptar(almacen.v_direccion0)
 Line Input #1, almacen.v_direccion1
 agenda.direccion1.Add escriptar.funcion_desescriptar(almacen.v_direccion1)
 Line Input #1, almacen.v_fecharegistro
 agenda.fecharegistro.Add escriptar.funcion_desescriptar(almacen.v_fecharegistro)
 Line Input #1, almacen.v_pais
 agenda.pais.Add escriptar.funcion_desescriptar(almacen.v_pais)
 Line Input #1, almacen.v_departamento
 agenda.departamento.Add escriptar.funcion_desescriptar(almacen.v_departamento)
 Line Input #1, almacen.v_ciudad
 agenda.ciudad.Add escriptar.funcion_desescriptar(almacen.v_ciudad)
 Line Input #1, almacen.v_calle0
 agenda.calle0.Add escriptar.funcion_desescriptar(almacen.v_calle0)
 Line Input #1, almacen.v_calle1
 agenda.calle1.Add escriptar.funcion_desescriptar(almacen.v_calle1)
 Line Input #1, almacen.v_calle2
 agenda.calle2.Add escriptar.funcion_desescriptar(almacen.v_calle2)
 Line Input #1, almacen.v_email0
 agenda.email0.Add escriptar.funcion_desescriptar(almacen.v_email0)
 Line Input #1, almacen.v_email1
 agenda.email1.Add escriptar.funcion_desescriptar(almacen.v_email1)
 Line Input #1, almacen.v_email2
 agenda.email2.Add escriptar.funcion_desescriptar(almacen.v_email2)
 Line Input #1, almacen.v_edad
 agenda.edad.Add escriptar.funcion_desescriptar(almacen.v_edad)
 Line Input #1, almacen.v_fn
 agenda.fn.Add escriptar.funcion_desescriptar(almacen.v_fn)
 Line Input #1, almacen.v_facebook0
 agenda.facebook0.Add escriptar.funcion_desescriptar(almacen.v_facebook0)
 Line Input #1, almacen.v_facebook1
 agenda.facebook1.Add escriptar.funcion_desescriptar(almacen.v_facebook1)
 Line Input #1, almacen.v_facebook2
 agenda.facebook2.Add escriptar.funcion_desescriptar(almacen.v_facebook2)
 Line Input #1, almacen.v_tuiter0
 agenda.tuiter0.Add escriptar.funcion_desescriptar(almacen.v_tuiter0)
 Line Input #1, almacen.v_tuiter1
 agenda.tuiter1.Add escriptar.funcion_desescriptar(almacen.v_tuiter1)
 Line Input #1, almacen.v_tuiter2
 agenda.tuiter2.Add escriptar.funcion_desescriptar(almacen.v_tuiter2)
 Line Input #1, almacen.v_nCasa
 agenda.nCasa.Add escriptar.funcion_desescriptar(almacen.v_nCasa)
 Line Input #1, almacen.v_Ecivil
 agenda.Ecivil.Add escriptar.funcion_desescriptar(almacen.v_Ecivil)
 End With
 Loop
 Close #1
no_se:
End Sub

Public Sub Guardar()
 modoEscrpto ' Escripto Bytes
 Open "bd.opxl" For Output As 1
 Dim i As Long
 With ModDatos.agenda
 For i = 1 To agenda.nombre.Count
 Print #1, escriptar.funcion_escriptar(.nombre.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.nombrex.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.apellidop.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.apellidom.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.telefono0.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.telefono1.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.hora.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.celular0.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.celular1.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.ci.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.direccion0.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.direccion1.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.fecharegistro.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.pais.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.departamento.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.ciudad.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.calle0.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.calle1.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.calle2.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.email0.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.email1.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.email2.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.edad.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.fn.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.facebook0.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.facebook1.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.facebook2.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.tuiter0.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.tuiter1.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.tuiter2.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.nCasa.Item(i).Key)
 Print #1, escriptar.funcion_escriptar(.Ecivil.Item(i).Key)
 Next
 End With
 Close #1
End Sub

Public Sub modoEscrpto()
 escriptar.variable_desescriptar_bytes = Mod_Funciones_conByts.escript_121
 escriptar.variable_escriptar_bytes = Mod_Funciones_conByts.escript_121
End Sub

