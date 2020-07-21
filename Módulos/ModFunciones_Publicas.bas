Attribute VB_Name = "ModFunciones_Publicas"
'***************************************************************************
'* Open Source
'* System Application Software - Funcines virtuales de Búsqueda
'* Módulo ModFunciones_Publicas de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************
Public Function busqueda_virtual(ByVal v_Busqueda As Byte)
On Error GoTo no_se
 Dim busco As Boolean
 Dim inteligente As Boolean
 Dim e As Long
        For e = 1 To agenda.nombre.Count
       Select Case (v_Busqueda)
      
              Case (0)
                   If (ModDatos.agenda.nombre.Item(e).Key = frmbusqueda.bdato.Text) Then 'Nombre
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (1)
                  If (ModDatos.agenda.nombrex.Item(e).Key = frmbusqueda.bdato.Text) Then 'Seg. Nombre
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (2)
               If (ModDatos.agenda.apellidom.Item(e).Key = frmbusqueda.bdato.Text) Then 'Apellido M
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (3)
              If (ModDatos.agenda.apellidop.Item(e).Key = frmbusqueda.bdato.Text) Then 'Apellido P
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (4)
             If (ModDatos.agenda.telefono0.Item(e).Key = frmbusqueda.bdato.Text) Then 'Telefono 0
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (5)
              If (ModDatos.agenda.telefono1.Item(e).Key = frmbusqueda.bdato.Text) Then 'Telefono 1
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (6)
              If (ModDatos.agenda.hora.Item(e).Key = frmbusqueda.bdato.Text) Then 'Telefono 2
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (7)
               If (ModDatos.agenda.celular0.Item(e).Key = frmbusqueda.bdato.Text) Then 'Celular  0
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (8)
               If (ModDatos.agenda.celular1.Item(e).Key = frmbusqueda.bdato.Text) Then 'Celular  1
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (9)
               If (ModDatos.agenda.ci.Item(e).Key = frmbusqueda.bdato.Text) Then 'CI
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (10)
               If (ModDatos.agenda.direccion0.Item(e).Key = frmbusqueda.bdato.Text) Then 'Direccion 0
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (11)
               If (ModDatos.agenda.direccion1.Item(e).Key = frmbusqueda.bdato.Text) Then 'Direccion 1
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (12)
              If (ModDatos.agenda.fecharegistro.Item(e).Key = frmbusqueda.bdato.Text) Then 'Direccion 2
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (13)
              If (ModDatos.agenda.pais.Item(e).Key = frmbusqueda.bdato.Text) Then 'Paìs
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (14)
              If (ModDatos.agenda.departamento.Item(e).Key = frmbusqueda.bdato.Text) Then 'Departamento
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (15)
               If (ModDatos.agenda.ciudad.Item(e).Key = frmbusqueda.bdato.Text) Then 'Ciudad
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (16)
               If (ModDatos.agenda.calle0.Item(e).Key = frmbusqueda.bdato.Text) Then 'Calle    0
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (17)
               If (ModDatos.agenda.calle1.Item(e).Key = frmbusqueda.bdato.Text) Then 'Calle    1
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (18)
              If (ModDatos.agenda.calle2.Item(e).Key = frmbusqueda.bdato.Text) Then ' Calle    2
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (19)
              If (ModDatos.agenda.email0.Item(e).Key = frmbusqueda.bdato.Text) Then 'Email 0
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (20)
               If (ModDatos.agenda.email1.Item(e).Key = frmbusqueda.bdato.Text) Then 'Email    1
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (21)
              If (ModDatos.agenda.email2.Item(e).Key = frmbusqueda.bdato.Text) Then 'Email 2
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (22)
               If (ModDatos.agenda.edad.Item(e).Key = frmbusqueda.bdato.Text) Then 'Edad
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (23)
              If (ModDatos.agenda.fn.Item(e).Key = frmbusqueda.bdato.Text) Then 'Fecha Nacimiento
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (24)
               If (ModDatos.agenda.facebook0.Item(e).Key = frmbusqueda.bdato.Text) Then 'Facebook 0
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (25)
             If (ModDatos.agenda.facebook1.Item(e).Key = frmbusqueda.bdato.Text) Then 'Facebook 1
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (26)
               If (ModDatos.agenda.facebook2.Item(e).Key = frmbusqueda.bdato.Text) Then 'Facebook 2
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (27)
               If (ModDatos.agenda.tuiter0.Item(e).Key = frmbusqueda.bdato.Text) Then 'Tuiter   0
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (28)
               If (ModDatos.agenda.tuiter1.Item(e).Key = frmbusqueda.bdato.Text) Then 'Tuiter   1
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (29)
              If (ModDatos.agenda.tuiter2.Item(e).Key = frmbusqueda.bdato.Text) Then 'Tuiter   2
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (30)
              If (ModDatos.agenda.nCasa.Item(e).Key = frmbusqueda.bdato.Text) Then 'Numero de Casa
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                   End If
              Case (31)
               If (ModDatos.agenda.Ecivil.Item(e).Key = frmbusqueda.bdato.Text) Then 'Estado Sibil
                       busco = True ' encontro dato
                       Else
                       busco = False ' no encontro dato
                       End If
End Select

'este código if regresa el selector a el estado original del listbox
If busco = True Then
   frmbusqueda.List1.ListIndex = -1
End If

Dim elemento_buscado As String
    elemento_buscado = "Indice: " & v_Busqueda & " etiqueta: " & frmbusqueda.List1.List(v_Busqueda) & " se encontro: " & busco
Select Case (busco)
      
       Case (True)
                With agenda
                 frmModificar.txtnombre(0).Text = .nombre.Item(e).Key
                 frmModificar.txtnombre(1).Text = .nombrex.Item(e).Key
                 frmModificar.txtnombre(2).Text = .apellidom.Item(e).Key
                 frmModificar.txtnombre(3).Text = .apellidop.Item(e).Key
                 frmModificar.txtnombre(4).Text = .telefono0.Item(e).Key
                 frmModificar.txtnombre(5).Text = .telefono1.Item(e).Key
                 frmModificar.DTPicker2.Value = .hora.Item(e).Key
                 frmModificar.txtnombre(7).Text = .celular0.Item(e).Key
                 frmModificar.txtnombre(15).Text = .celular1.Item(e).Key
                 frmModificar.txtnombre(14).Text = .ci.Item(e).Key
                 frmModificar.txtnombre(13).Text = .direccion0.Item(e).Key
                 frmModificar.txtnombre(12).Text = .direccion1.Item(e).Key
                 frmModificar.txtnombre(11).Text = .fecharegistro.Item(e).Key
                 frmModificar.txtnombre(10).Text = .pais.Item(e).Key
                 frmModificar.txtnombre(9).Text = .departamento.Item(e).Key
                 frmModificar.txtnombre(8).Text = .ciudad.Item(e).Key
                 frmModificar.txtnombre(31).Text = .calle0.Item(e).Key
                 frmModificar.txtnombre(30).Text = .calle1.Item(e).Key
                 frmModificar.txtnombre(29).Text = .calle2.Item(e).Key
                 frmModificar.txtnombre(28).Text = .email0.Item(e).Key
                 frmModificar.txtnombre(27).Text = .email1.Item(e).Key
                 frmModificar.txtnombre(26).Text = .email2.Item(e).Key
                 frmModificar.txtnombre(25).Text = .edad.Item(e).Key
                 frmModificar.txtnombre(24).Text = .fn.Item(e).Key
                 frmModificar.txtnombre(16).Text = .facebook0.Item(e).Key
                 frmModificar.txtnombre(17).Text = .facebook1.Item(e).Key
                 frmModificar.txtnombre(18).Text = .facebook2.Item(e).Key
                 frmModificar.txtnombre(19).Text = .tuiter0.Item(e).Key
                 frmModificar.txtnombre(20).Text = .tuiter1.Item(e).Key
                 frmModificar.txtnombre(21).Text = .tuiter2.Item(e).Key
                 frmModificar.txtnombre(22).Text = .nCasa.Item(e).Key
                 frmModificar.txtnombre(23).Text = .Ecivil.Item(e).Key
                 frmModificar.Show 1
                 frmbusqueda.Timer1.Enabled = False
            End With
             ' MsgBox elemento_buscado
       Case (False)
         
           If (frmbusqueda.List1.ListIndex = 31) Then
                 frmbusqueda.Timer1.Enabled = False
                 'frmbusqueda.Check1 = 0
                 inteligente = True
                 frmbusqueda.List1.ListIndex = -1
                 End If
                 If (inteligente = True) Then
                 inteligente = False
                 If frmModificar.Visible = False Then
                    MsgBox "No se Encontraron Datos en BD.", vbInformation, nombre_programa
                 End If
             End If
End Select
Next
frmbusqueda.List1.ListIndex = frmbusqueda.List1.ListIndex + 1
no_se:
End Function

Public Function virtual_ci(ByVal ci As String)
       Dim e As Long
       On Error GoTo no_se
       For e = 1 To agenda.ci.Count
           If (ModDatos.agenda.ci.Item(e).Key = ci) Then 'CI
                      ' frmdatos.Command1(0).Enabled = False ' para que no pueda crear el registro
                      ' MsgBox "Existe la CI por lo tanto no se podra Crear Duplicados de esta Identidad"
                       Else
                       frmdatos.Command1(0).Enabled = True
             End If
        Next
no_se:
End Function

