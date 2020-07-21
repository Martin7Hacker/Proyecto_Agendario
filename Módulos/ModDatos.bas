Attribute VB_Name = "ModDatos"
'***************************************************************************
'* Open Source
'* System Application Software
'* Módulo ModDatos de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************
Public ag As Datos
Public Const nombre_programa = "Agendario Express v1.0"
Public comparador As Long
Public agenda As agendax
Public almacen As almacenx
Public seguridad As datsg
Public y_seguridad As y_datsg

Private Type agendax
            nombre                As New Datos
            nombrex               As New Datos
            apellidop             As New Datos
            apellidom             As New Datos
            telefono0             As New Datos
            telefono1             As New Datos
            hora                  As New Datos
            celular0              As New Datos
            celular1              As New Datos
            ci                    As New Datos
            direccion0            As New Datos
            direccion1            As New Datos
            fecharegistro         As New Datos
            pais                  As New Datos
            departamento          As New Datos
            ciudad                As New Datos
            calle0                As New Datos
            calle1                As New Datos
            calle2                As New Datos
            email0                As New Datos
            email1                As New Datos
            email2                As New Datos
            edad                  As New Datos
            fn                    As New Datos
            facebook0             As New Datos
            facebook1             As New Datos
            facebook2             As New Datos
            tuiter0               As New Datos
            tuiter1               As New Datos
            tuiter2               As New Datos
            nCasa                 As New Datos
            Ecivil                As New Datos
            
End Type

Private Type almacenx
            v_nombre                As String
            v_nombrex               As String
            v_apellidop             As String
            v_apellidom             As String
            v_telefono0             As String
            v_telefono1             As String
            v_hora                  As String
            v_celular0              As String
            v_celular1              As String
            v_ci                    As String
            v_direccion0            As String
            v_direccion1            As String
            v_fecharegistro         As String
            v_pais                  As String
            v_departamento          As String
            v_ciudad                As String
            v_calle0                As String
            v_calle1                As String
            v_calle2                As String
            v_email0                As String
            v_email1                As String
            v_email2                As String
            v_edad                  As String
            v_fn                    As String
            v_facebook0             As String
            v_facebook1             As String
            v_facebook2             As String
            v_tuiter0               As String
            v_tuiter1               As String
            v_tuiter2               As String
            v_nCasa                 As String
            v_Ecivil                As String
End Type

Private Type datsg
            x_inicio                As String
            x_modific               As String
            x_creacion              As String
            x_busqueda              As String
            x_poderver              As String
            x_iniciodel             As String
            x_elimnarTodo           As String
            x_eliminarSeleccionado  As String
            x_poderExportar         As String
            x_poderImprimir         As String
            x_crearCopiaSeguridad   As String
          
End Type

Private Type y_datsg
            y_inicio                As String
            y_modific               As String
            y_creacion              As String
            y_busqueda              As String
            y_poderver              As String
            y_iniciodel             As String
            y_elimnarTodo           As String
            y_eliminarSeleccionado  As String
            y_poderExportar         As String
            y_poderImprimir         As String
            y_crearCopiaSeguridad   As String
            
          
End Type

Public Sub abrir_archivo()
 mod_Guardar_Abrir.Abrir
 If Not frmvisualizar.ListView1.ListItems.Count = 0 Then
 frmvisualizar.visualizar_todos_los_registros
 frmvisualizar.elimino_todo
 Else
 frmvisualizar.visualizar_todos_los_registros
 End If
End Sub

Public Sub guardar_archivo()
 mod_Guardar_Abrir.Guardar
End Sub

