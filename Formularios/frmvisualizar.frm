VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmvisualizar 
   BackColor       =   &H00800080&
   Caption         =   "Planilla Virtual - Virtual Sheet"
   ClientHeight    =   8190
   ClientLeft      =   120
   ClientTop       =   -1770
   ClientWidth     =   12900
   Icon            =   "frmvisualizar.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8190
   ScaleWidth      =   12900
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4455
      ScaleWidth      =   375
      TabIndex        =   8
      Top             =   0
      Width           =   375
      Begin VB.PictureBox picabc 
         BackColor       =   &H00800080&
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   375
         TabIndex        =   9
         Top             =   0
         Width           =   375
         Begin VB.CommandButton cmdletra 
            Caption         =   "A"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   375
         End
      End
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   4815
      Left            =   360
      Max             =   -1
      Min             =   -10
      TabIndex        =   11
      Top             =   0
      Value           =   -1
      Width           =   255
   End
   Begin VB.CommandButton cmdmasmenos 
      Height          =   375
      Index           =   0
      Left            =   50
      Picture         =   "frmvisualizar.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   375
   End
   Begin VB.CommandButton cmdmasmenos 
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2000
      Picture         =   "frmvisualizar.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   2415
      TabIndex        =   13
      Top             =   0
      Width           =   2415
   End
   Begin VB.CommandButton cmdabc 
      Caption         =   "abc"
      Height          =   255
      Left            =   1800
      TabIndex        =   12
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton cmdcod 
      Caption         =   "inf"
      Height          =   255
      Left            =   0
      Picture         =   "frmvisualizar.frx":265E
      TabIndex        =   7
      Top             =   2160
      Width           =   320
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7815
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   7125
      _ExtentX        =   12568
      _ExtentY        =   13785
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   8388736
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
      Picture         =   "frmvisualizar.frx":3328
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   9075
      Left            =   2400
      Max             =   -1
      Min             =   -10
      TabIndex        =   6
      Top             =   0
      Value           =   -1
      Width           =   255
   End
   Begin VB.PictureBox picuteMesShop 
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      Height          =   9135
      Left            =   0
      Picture         =   "frmvisualizar.frx":627DA
      ScaleHeight     =   9135
      ScaleWidth      =   2775
      TabIndex        =   3
      Top             =   0
      Width           =   2775
      Begin VB.PictureBox panel1 
         BackColor       =   &H00800080&
         BorderStyle     =   0  'None
         Height          =   3375
         Left            =   0
         ScaleHeight     =   3375
         ScaleWidth      =   2535
         TabIndex        =   4
         Top             =   0
         Width           =   2535
         Begin MSComCtl2.MonthView meses 
            Height          =   2460
            Index           =   0
            Left            =   0
            TabIndex        =   5
            Top             =   0
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   4339
            _Version        =   393216
            ForeColor       =   8388736
            BackColor       =   8388736
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MonthBackColor  =   16777215
            ShowToday       =   0   'False
            StartOfWeek     =   49676290
            TitleBackColor  =   8388736
            TitleForeColor  =   16777215
            CurrentDate     =   41776
         End
      End
   End
   Begin VB.Menu Archivo 
      Caption         =   "&Archivo - &File"
      Begin VB.Menu Agendar 
         Caption         =   "&Agendar - &Schedule"
         Shortcut        =   {F1}
      End
      Begin VB.Menu Buscar 
         Caption         =   "&Búscar - &Search"
         Shortcut        =   {F2}
      End
      Begin VB.Menu Modificarx 
         Caption         =   "&Modificar - &Modify"
         Shortcut        =   {F3}
      End
      Begin VB.Menu Eliminar 
         Caption         =   "&Eliminar - &Remove"
         Shortcut        =   {F4}
      End
      Begin VB.Menu PlanillaVirtual 
         Caption         =   "&Planilla Virtual - &Virtual Sheet"
         Shortcut        =   {F5}
      End
      Begin VB.Menu Exportar 
         Caption         =   "&Exportar - &To export"
         Shortcut        =   {F6}
      End
      Begin VB.Menu Imprimirx 
         Caption         =   "&Imprimir - &Print"
         Shortcut        =   {F7}
      End
      Begin VB.Menu Salir 
         Caption         =   "&Salir - &Exit"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu explorarmartinsoft 
      Caption         =   "&Explorar Martinsoft - &explore &Martinsoft"
   End
   Begin VB.Menu clima 
      Caption         =   "&Clima - &Weather"
   End
   Begin VB.Menu ExportaroImprimir 
      Caption         =   "&Exportar o Imprimir - &Export or &Print"
      Begin VB.Menu Exportarcomo 
         Caption         =   "&Exportar como... - &export &as"
      End
      Begin VB.Menu Imprimir 
         Caption         =   "&Imprimir - &Print"
      End
   End
   Begin VB.Menu Datos 
      Caption         =   "&Datos - &Data"
      Begin VB.Menu ver 
         Caption         =   "&Ver - &view"
         Begin VB.Menu VerPlanillaVirtual 
            Caption         =   "&Ver Planilla Virtual - &view Virtual Sheet"
         End
         Begin VB.Menu ocultarplanillavirtual 
            Caption         =   "&Ocultar Planilla Virtual - &Hide Virtual &Sheet"
         End
         Begin VB.Menu Calendario 
            Caption         =   "&Calendario - &Calendar"
         End
         Begin VB.Menu VizualizartodoslosRegistros 
            Caption         =   "Visualizar Registros - &display &Records"
         End
         Begin VB.Menu Verregistrosconfiltro 
            Caption         =   "&Ver registros con filtro - &View records &filter"
         End
      End
      Begin VB.Menu ModificaciondeDatos 
         Caption         =   "Modificación de Datos - &Data &Modification"
         Begin VB.Menu Modificar 
            Caption         =   "&Modificar - &Modify"
         End
         Begin VB.Menu perzonalizaratos 
            Caption         =   "&Perzonalizar Datos - &Customise &Details"
         End
      End
      Begin VB.Menu formadeEliminar 
         Caption         =   "&Forma de Eliminar - &Delete &form"
         Begin VB.Menu EliminarTodo 
            Caption         =   "&Elimnar Todo - &Delete &all"
         End
         Begin VB.Menu Eliminar_Selecionadox 
            Caption         =   "&Eliminar Seleciónado - remove Selected"
         End
      End
   End
   Begin VB.Menu Seguridad 
      Caption         =   "&Seguridad - &Security"
      Begin VB.Menu contraceña 
         Caption         =   "&Contraceña - &Password"
      End
      Begin VB.Menu crearcopiadeseguridad 
         Caption         =   "&Crear copia de Seguridad - &Create &Backup"
      End
   End
   Begin VB.Menu Escritorio 
      Caption         =   "&Escritorio - &Desktop"
      Visible         =   0   'False
      Begin VB.Menu definirfondo 
         Caption         =   "Definir Fondo de Escritorio - &Define &wallpaper"
      End
   End
   Begin VB.Menu Ayuda 
      Caption         =   "&Ayuda - Help"
      Begin VB.Menu donaci 
         Caption         =   "&Donación $ para Software Agendario Express v1.0 - Donation $ Agendario Express Software v1.0"
         Shortcut        =   {F11}
      End
      Begin VB.Menu AyudadelPrograma 
         Caption         =   "&Ayuda del Programa - &Help &Program"
         Shortcut        =   {F12}
      End
      Begin VB.Menu AcercadeAgendario 
         Caption         =   "&Acerca de Agendario - &About &Agendario"
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu definidos 
      Caption         =   "&definidos - &defined"
      Visible         =   0   'False
      Begin VB.Menu MostrartodoslosMeses 
         Caption         =   "&Mostrar todos los &Meses - &Show all &Months"
      End
      Begin VB.Menu esx 
         Caption         =   "-"
      End
      Begin VB.Menu SolodefinidosActuales 
         Caption         =   "&Solo definidos &Actuales - &Only defined &Current"
      End
   End
   Begin VB.Menu abcx 
      Caption         =   "abcx"
      Visible         =   0   'False
      Begin VB.Menu Verabecedario 
         Caption         =   "&Ver &abecedario - &View &alphabet"
      End
      Begin VB.Menu espxg 
         Caption         =   "-"
      End
      Begin VB.Menu Ocultarabecedario 
         Caption         =   "&Ocultar &abecedario - &Hide &alphabet "
      End
   End
End
Attribute VB_Name = "frmvisualizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Open Source
'* System Application Software
'* Programa frmvisualizar de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************
Dim proceso_x As Boolean: Dim elemento As ListItem: Dim borro As Long

Private Sub agregar_registros()
 With ListView1
 ListView1.ColumnHeaders.Clear
 Open "etiquetas.opxl" For Input As 1
 Dim etiqueta As String
 Do While Not EOF(1)
 Line Input #1, etiqueta
 .ColumnHeaders.Add , , escriptar.funcion_desescriptar(etiqueta)
 Loop
 Close #1
 End With
End Sub

Public Sub cmdeliminartodo_Click()
 Select Case MsgBox("Quieres eliminar todos los registros de la Planilla Virtual y borrar los datos alojados en la base de datos de " & _
 nombre_programa, vbExclamation + vbYesNo, nombre_programa)
 Case (vbYes)
 If ListView1.ListItems.Count = 0 Then
 MsgBox "No existen registros Cargados por lo tanto si no existe no se eliminara nada", vbInformation, ModDatos.nombre_programa
 Else
 funcion_elimniar_todo ' funcion prar eliminar todos los datos de Agendario
 End If
 End Select
End Sub

Public Sub cmdModificar_Click()
 If ListView1.ListItems.Count = 0 Then
 comarador = 0
 End If
 If Not (comparador = 0) Then
 If ModDatos.y_seguridad.y_modific = "" Then
 frmModificar.Show 1
 Else
 frmModificarDatos.Show 1
 End If
 Else
 Mod_Funciones_conByts.desoprimr_boton 6
 MsgBox "Para poder modificar el registro tienes que seleciónar un registro", _
 vbInformation, nombre_programa
 End If
End Sub

Public Sub cmdVisualizar_Click()
 comparador = 0 ' establesco el selecionador en nada selecionado
 visualizar_todos_los_registros
End Sub

Private Sub acercadeagendario_Click()
 frmAcercade.Show 1
End Sub

Private Sub Agendar_Click()
 MDIPrincipal.crearNuevaHoja
End Sub

Private Sub AyudadelPrograma_Click()
 Mod_Funciones_conByts.openAyuda Me
End Sub

Private Sub buscar_Click()
 MDIPrincipal.buscarEnBD
End Sub

Private Sub Calendario_Click()
 frmcalendario.Show 1
End Sub

Private Sub verabc(ByVal estado As Boolean)
 picabc.Visible = estado
 Picture2.Visible = estado
 VScroll2.Visible = estado
End Sub

Private Sub clima_Click()
 Mod_Funciones_conByts.consultartiempo Me
End Sub

Private Sub cmdabc_Click()
 PopupMenu abcx
End Sub

Private Sub cmdcod_Click()
 PopupMenu definidos
End Sub

Private Sub cmdletra_Click(Index As Integer)
 On Error GoTo nose
 ListView1.SetFocus
nose:
End Sub

Private Sub cmdmasmenos_Click(Index As Integer)
 Dim dias_m, mes_n As Byte
 Dim mese_s, anio As Integer
 Select Case Index
 Case (0)
 For mese_s = 0 To 11
  mes_n = mes_n + 1
  anio = meses(mese_s).Year + 1
  meses(mese_s).Value = "01/" & "" & mes_n & "" & " / " & "" & anio & ""
  Next mese_s
  despinarTodoslosMeses
  Case (1)
 For mese_s = 0 To 11
  mes_n = mes_n + 1
  anio = meses(mese_s).Year - 1
  If anio = 1752 Then
  MsgBox "el Año minimo es 1753", vbInformation, nombre_programa
  Exit Sub
  Else
  meses(mese_s).Value = "01/" & "" & mes_n & "" & " / " & "" & anio & ""
  End If
  Next mese_s
  despinarTodoslosMeses
End Select
End Sub

Private Sub contraceña_Click()
 If y_seguridad.y_iniciodel = "" Then
 frmseguridad.Show 1
 Else
 frmseguridadx.Show 1
 End If
End Sub

Private Sub crearcopiadeseguridad_Click()
 If y_seguridad.y_crearCopiaSeguridad = "" Then
 frmcopiaSeguridad.Show 1
 Else
 frmcrearcopiadeSeguridad.Show 1
 End If
End Sub

Private Sub definirfondo_Click()
 frmfondoPantalla.Show 1
End Sub

Private Sub donaci_Click()
 Mod_Funciones_conByts.oprimir_boton 19
 Select Case MsgBox("Quieres Realizar una Donacion para Ayudar al Proyecto Agendario Express v1.0", _
 vbYesNo + vbInformation, nombre_programa)
 Case (vbYes)
 Mod_Funciones_conByts.desoprimr_boton 19
 Mod_Funciones_conByts.AbrirWeb Me, "http://adf.ly/1TJlyy"
 Case (vbNo)
 Mod_Funciones_conByts.desoprimr_boton 19
 End Select
End Sub

Private Sub Eliminar_Click()
 MDIPrincipal.elimnarSelecionado
End Sub

Private Sub Eliminar_Selecionadox_Click()
 If y_seguridad.y_eliminarSeleccionado = "" Then
 eliminarselecionado_Click
 Else
 frmElimnarSelecionado.Show 1
 End If
End Sub

Public Sub eliminarselecionado_Click()
 If ListView1.ListItems.Count = 0 Then
 comarador = 0
 End If
 If Not (comparador = 0) Then
 Select Case MsgBox("Quieres eliminar el registro seleccionado en la Planilla Virtual y borrar los datos alojados en la base de datos de " _
 & nombre_programa, vbExclamation + vbYesNo)
 Case (vbYes)
 Eliminar_Selecionado ' procedimiento que elimina el registro selecionado
 desoprimr_boton (8)
 Case (vbNo)
 desoprimr_boton (8)
 End Select
 Else
 MsgBox "Seleccióne el Registro que quiere eliminar", _
 vbInformation, nombre_programa
 End If
End Sub

Private Sub EliminarTodo_Click()
 If y_seguridad.y_elimnarTodo = "" Then
 cmdeliminartodo_Click
 Else
 frmEliminarTodo.Show 1
 End If
End Sub

Private Sub explorarmartinsoft_Click()
 frmconsultas.Show 1
End Sub

Private Sub Exportar_Click()
 MDIPrincipal.ExportarArchivo
End Sub

Private Sub Exportarcomo_Click()
 On Error GoTo nose
 If y_seguridad.y_poderExportar = "" Then
 frmexportar_datos.Show 1
 Else
 frmExportarDatos.Show 1
 End If
nose:
End Sub

Private Sub Form_Load()
 crear_meses
 cmdmeses_Click 12
 comparador = 0
 agregar_registros
 With ListView1
 .LabelEdit = lvwManual
 .AllowColumnReorder = True
 .FullRowSelect = True
 End With
 acoplar
 Mod_Funciones_conByts.oprimir_boton 10
 abc
 verabc False
 If Not (y_seguridad.y_poderver = "") Then
 frmvisualizar.ListView1.Visible = False
 End If
 frmvisualizar.ListView1.Picture = LoadPicture("img\lista.bmp")
End Sub

Private Sub acoplar()
 On Error GoTo nose
 With ListView1
 .Width = Me.Width - 2750
 .Height = Me.Height - 500
 .Top = 0
 .Left = 2650
 VScroll1.Height = Me.Height - 500
 picuteMesShop.Height = Me.Height - 500
 Picture2.Height = Me.Height - 500
 VScroll2.Height = Me.Height - 500
 End With
nose:
End Sub

Public Sub visualizar_todos_los_registros()
 Dim r As Long
 On Error GoTo no_se
 ListView1.ListItems.Clear
 Dim i As Long
 For i = 1 To agenda.nombre.Count
 If Not (agenda.nombre.Count = 1) Then
 With frmcargando
 .Show
 .ProgressBar1.Max = agenda.nombre.Count
 .ProgressBar1.Min = 1
 End With
 End If
 Set elemento = ListView1.ListItems.Add(, , agenda.nombre.Item(i).Key)
 frmcargando.ProgressBar1.Value = i
 r = i
 frmcargando.Caption = "Cargando datos Espere Por Favor ..." & " " & "Archivo :" & " " & r
 elemento.SubItems(1) = agenda.nombrex.Item(i).Key
 elemento.SubItems(2) = agenda.apellidom.Item(i).Key
 elemento.SubItems(3) = agenda.apellidop.Item(i).Key
 elemento.SubItems(4) = agenda.telefono0.Item(i).Key
 elemento.SubItems(5) = agenda.telefono1.Item(i).Key
 elemento.SubItems(6) = agenda.hora.Item(i).Key
 elemento.SubItems(7) = agenda.celular0.Item(i).Key
 elemento.SubItems(8) = agenda.celular1.Item(i).Key
 elemento.SubItems(9) = agenda.ci.Item(i).Key
 elemento.SubItems(10) = agenda.direccion0.Item(i).Key
 elemento.SubItems(11) = agenda.direccion1.Item(i).Key
 elemento.SubItems(12) = agenda.fecharegistro.Item(i).Key
 elemento.SubItems(13) = agenda.pais.Item(i).Key
 elemento.SubItems(14) = agenda.departamento.Item(i).Key
 elemento.SubItems(15) = agenda.ciudad.Item(i).Key
 elemento.SubItems(16) = agenda.calle0.Item(i).Key
 elemento.SubItems(17) = agenda.calle1.Item(i).Key
 elemento.SubItems(18) = agenda.calle2.Item(i).Key
 elemento.SubItems(19) = agenda.email0.Item(i).Key
 elemento.SubItems(20) = agenda.email1.Item(i).Key
 elemento.SubItems(21) = agenda.email2.Item(i).Key
 elemento.SubItems(22) = agenda.edad.Item(i).Key
 elemento.SubItems(23) = agenda.fn.Item(i).Key
 elemento.SubItems(24) = agenda.facebook0.Item(i).Key
 elemento.SubItems(25) = agenda.facebook1.Item(i).Key
 elemento.SubItems(26) = agenda.facebook2.Item(i).Key
 elemento.SubItems(27) = agenda.tuiter0.Item(i).Key
 elemento.SubItems(28) = agenda.tuiter1.Item(i).Key
 elemento.SubItems(29) = agenda.tuiter2.Item(i).Key
 elemento.SubItems(30) = agenda.nCasa.Item(i).Key
 elemento.SubItems(31) = agenda.Ecivil.Item(i).Key
 Next
 Unload frmcargando
no_se:
End Sub

Private Sub Eliminar_Selecionado()
 On Error GoTo nose
 With ListView1
 If Not (.ListItems.Count = 0) Or .SelectedItem.Index = 0 Then
 .ListItems.Remove (.SelectedItem.Index)
 End If
 End With
 With agenda
 .nombre.Remove (borro)
 .nombrex.Remove (borro)
 .apellidom.Remove (borro)
 .apellidop.Remove (borro)
 .telefono0.Remove (borro)
 .telefono1.Remove (borro)
 .hora.Remove (borro)
 .celular0.Remove (borro)
 .celular1.Remove (borro)
 .ci.Remove (borro)
 .direccion0.Remove (borro)
 .direccion1.Remove (borro)
 .fecharegistro.Remove (borro)
 .pais.Remove (borro)
 .departamento.Remove (borro)
 .ciudad.Remove (borro)
 .calle0.Remove (borro)
 .calle1.Remove (borro)
 .calle2.Remove (borro)
 .email0.Remove (borro)
 .email1.Remove (borro)
 .email2.Remove (borro)
 .edad.Remove (borro)
 .fn.Remove (borro)
 .facebook0.Remove (borro)
 .facebook1.Remove (borro)
 .facebook2.Remove (borro)
 .tuiter0.Remove (borro)
 .tuiter1.Remove (borro)
 .tuiter2.Remove (borro)
 .nCasa.Remove (borro)
 .Ecivil.Remove (borro)
 End With
 comparador = 0
 ModDatos.guardar_archivo
nose:
End Sub

Private Sub Form_Resize()
 acoplar
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Mod_Funciones_conByts.desoprimr_boton 10
End Sub

Private Sub imprimir_Click()
 On Error GoTo no_se
 If y_seguridad.y_poderImprimir = "" Then
 frmImprimir.Show 1
 Else
 frmpoderImprimir.Show 1
 End If
no_se:
End Sub

Private Sub Imprimirx_Click()
 MDIPrincipal.Inprimir
End Sub

Private Sub ListView1_Click()
 On Error GoTo no_se
 comparador = ListView1.SelectedItem.Index
 borro = comparador
no_se:
End Sub

Private Sub funcion_elimniar_todo()
 Select Case MsgBox("¿ Queres eliminar todos los registro de " & _
 nombre_programa & " para no poder ver los registros  nunca más ?", _
 vbExclamation + vbYesNo, nombre_programa)
 Case (vbYes)
 On Error GoTo no_se
 With agenda
 Dim x As Long
 For x = 0 To agenda.nombre.Count
 .nombre.Limpiar
 .nombrex.Limpiar
 .apellidom.Limpiar
 .apellidop.Limpiar
 .telefono0.Limpiar
 .telefono1.Limpiar
 .hora.Limpiar
 .celular0.Limpiar
 .celular1.Limpiar
 .ci.Limpiar
 .direccion0.Limpiar
 .direccion1.Limpiar
 .fecharegistro.Limpiar
 .pais.Limpiar
 .departamento.Limpiar
 .ciudad.Limpiar
 .calle0.Limpiar
 .calle1.Limpiar
 .calle2.Limpiar
 .email0.Limpiar
 .email1.Limpiar
 .email2.Limpiar
 .edad.Limpiar
 .fn.Limpiar
 .facebook0.Limpiar
 .facebook1.Limpiar
 .facebook2.Limpiar
 .tuiter0.Limpiar
 .tuiter1.Limpiar
 .tuiter2.Limpiar
 .nCasa.Limpiar
 .Ecivil.Limpiar
 Next
 visualizar_todos_los_registros
 End With
no_se:
 End Select
 ModDatos.guardar_archivo
End Sub

Public Sub elimino_todo()
 With ModDatos.agenda
 Dim x As Long
 For x = 0 To agenda.nombre.Count
 .nombre.Limpiar
 .nombrex.Limpiar
 .apellidom.Limpiar
 .apellidop.Limpiar
 .telefono0.Limpiar
 .telefono1.Limpiar
 .hora.Limpiar
 .celular0.Limpiar
 .celular1.Limpiar
 .ci.Limpiar
 .direccion0.Limpiar
 .direccion1.Limpiar
 .fecharegistro.Limpiar
 .pais.Limpiar
 .departamento.Limpiar
 .ciudad.Limpiar
 .calle0.Limpiar
 .calle1.Limpiar
 .calle2.Limpiar
 .email0.Limpiar
 .email1.Limpiar
 .email2.Limpiar
 .edad.Limpiar
 .fn.Limpiar
 .facebook0.Limpiar
 .facebook1.Limpiar
 .facebook2.Limpiar
 .tuiter0.Limpiar
 .tuiter1.Limpiar
 .tuiter2.Limpiar
 .nCasa.Limpiar
 .Ecivil.Limpiar
 Next
 End With
End Sub

Private Sub ListView1_MouseDown(Button As Integer, Shift _
As Integer, x As Single, y As Single)
 If Button = vbRightButton Then
 PopupMenu Datos ' muestra un menú deslizable en pantalla
 End If
End Sub

Private Sub meses_DateClick(Index As Integer, ByVal DateClicked As Date)
 frmbusqueda.bdato.Text = meses(Index).Value
 frmbusqueda.Show 1
End Sub

Private Sub Modificar_Click()
 cmdModificar_Click
End Sub

Private Sub Modificarx_Click()
  frmvisualizar.cmdModificar_Click
End Sub

Private Sub MostrartodoslosMeses_Click()
 proceso_x = True
End Sub

Private Sub Ocultarabecedario_Click()
 verabc False
End Sub

Private Sub ocultarplanillavirtual_Click()
 ListView1.Visible = False
End Sub

Private Sub perzonalizaratos_Click()
 If y_seguridad.y_iniciodel = "" Then
 frmPerzonalizarDatos.Show 1
 Else
 frmsegurdiadPerzonalizar.Show 1
 End If
End Sub

Private Sub planillavirtual_Click()
 MDIPrincipal.MostrarPlanillaVirtual
End Sub

Private Sub salir_Click()
 Unload Me
End Sub

Private Sub SolodefinidosActuales_Click()
 proceso_x = False
End Sub

Private Sub Verabecedario_Click()
 verabc True
End Sub

Private Sub VerPlanillaVirtual_Click()
 If y_seguridad.y_poderver = "" Then
 ListView1.Visible = True
 Else
 frmpoderverx.Show 1
 End If
End Sub

Private Sub Verregistrosconfiltro_Click()
 frmfiltro.Show 1
End Sub

Private Sub VizualizartodoslosRegistros_Click()
 cmdVisualizar_Click
End Sub

Function Exportar_Excel(sFileName As String, ListView As ListView, _
Optional Progressbar As Progressbar, Optional SheetIndex As Integer = 1) _
As Boolean
 On Error GoTo error_Handler
 Dim obj_Excel, obj_Libro As Object: Dim iCol, iRow As Integer
 Set obj_Excel = CreateObject("Excel.Application")
 With obj_Excel
 Set obj_Libro = .Workbooks.Open(sFileName)
 End With
 With obj_Libro
 If Not Progressbar Is Nothing Then
 Progressbar.Max = ListView.ListItems.Count
 If Not Progressbar.Visible Then Progressbar.Visible = True
 End If
 With .Sheets(SheetIndex)
 For iRow = 1 To ListView.ListItems.Count
 iCol = 1
 .Cells(iRow, iCol) = ListView.ListItems.Item(iRow)
 For iCol = 1 To ListView.ColumnHeaders.Count - 1
 .Cells(iRow, iCol + 1) = ListView.ListItems(iRow).SubItems(iCol)
 Next
 If Not Progressbar Is Nothing Then
 Progressbar.Value = Progressbar.Value + 1
 End If
 Next
 End With
 End With
 obj_Excel.Visible = True
 Set obj_Libro = Nothing
 Set obj_Excel = Nothing
 Exportar_Excel = True
 If Not Progressbar Is Nothing Then
 Progressbar.Value = 0
 Progressbar.Visible = False
 End If
 Exit Function
error_Handler:
 Exportar_Excel = False
 MsgBox Err.Description, vbCritical
 On Error Resume Next
 Set obj_Libro = Nothing
 Set obj_Excel = Nothing
 Progressbar.Value = 0
End Function

Public Sub exportaraExel_Click()
 Dim ret As Boolean
 ret = Exportar_Excel(frmexportar_datos.Text2, ListView1, frmExportaraExel.ProgressBar1, 2)
 If ret Then
 MsgBox " Archivo Excel generado en: " & vbCrLf & frmexportar_datos.Text2 _
 , vbInformation
 End If
End Sub

Private Sub VScroll1_Change()
 panel1.Top = VScroll1.Value * 2280
 If proceso_x = True And VScroll1.Value = 0 Then
 VScroll1.Value = -8
 cmdmasmenos_Click 0
 despinarTodoslosMeses
 End If
 If proceso_x = True And VScroll1.Value = -9 Then
 VScroll1.Value = -1
 cmdmasmenos_Click 1
 despinarTodoslosMeses
 End If
End Sub

Private Sub despinarTodoslosMeses()
 On Error GoTo nose
 Dim dias_x As Byte: Dim anio_a, anio_c As Integer
 Dim ultimoDiaMes As String
 anio_a = Mid(Date, 7, 10)
 anio_c = meses(0).Year
 For dias_x = 0 To 11
 meses(dias_x).Font.Underline = False
 meses(dias_x).Font.Strikethrough = False
 If anio_a < meses(dias_x).Year Then
 meses(dias_x).Font.Underline = True
 meses(dias_x).Day = 1
 ElseIf anio_a > meses(dias_x).Year Then
 meses(dias_x).Font.Strikethrough = True
 ultimoDiaMes = DateSerial(Year(Now), meses(dias_x).Month + 1, 0)
 ultimoDiaMes = Mid(ultimoDiaMes, 1, 2)
 meses(dias_x).Day = ultimoDiaMes
 ElseIf anio_a = meses(dias_x).Year Then
 anioIgualaAnio
 End If
 Next dias_x
nose:
End Sub

Private Sub anioIgualaAnio()
 Dim dias As Byte: Dim ultimoDiaMes As String
 Dim anio As Integer
 For dias = 0 To 11
 meses(dias).Font.Underline = False
 meses(dias).Font.Strikethrough = False
 Next dias
 For dias = 0 To mesDelAnio - 1
 ultimoDiaMes = DateSerial(Year(Now), meses(dias).Month + 1, 0)
 ultimoDiaMes = Mid(ultimoDiaMes, 1, 2)
 meses(dias).Day = ultimoDiaMes
 meses(dias).Font.Strikethrough = True
 Next dias
 meses(mesDelAnio).Day = Day(Date)
 For dias = mesDelAnio + 1 To 11
 meses(dias).Day = 1
 meses(dias).Font.Underline = True
 Next dias
End Sub

Private Sub VScroll1_Scroll()
 VScroll1_Change
End Sub

Private Sub crear_meses()
 Dim meses_d As Byte
 For meses_d = 1 To 12
 l_meses = l_meses + 1
 Load meses(l_meses)
 meses(l_meses).Visible = True
 meses(l_meses).Top = 2280 * l_meses
 meses(l_meses).TitleBackColor = meses(0).TitleBackColor
 meses(l_meses).MonthBackColor = meses(0).MonthBackColor
 panel1.Height = 2280 * l_meses
 With VScroll1
 .Min = 0
 .Max = -l_meses + 3
 End With
 Next
 meses(0).Month = mvwJanuary   'enero
 meses(1).Month = mvwFebruary  'febrero
 meses(2).Month = mvwMarch     'marso
 meses(3).Month = mvwApril     'abril
 meses(4).Month = mvwMay       'mayo
 meses(5).Month = mvwJune      'junio
 meses(6).Month = mvwJuly      'julio
 meses(7).Month = mvwAugust    'agosto
 meses(8).Month = mvwSeptember 'septiembre
 meses(9).Month = mvwOctober   'octubre
 meses(10).Month = mvwNovember 'noviembre
 meses(11).Month = mvwDecember 'diciembre
End Sub

Private Function pintarMeses()
 pintarMeses = &H800080
End Function

Public Sub cmdmeses_Click(Index As Integer)
 Dim anio_x As Byte
 With VScroll1
 proceso_x = False
 Select Case Index
 Case 0: .Value = 0
 Case 1: .Value = -1
 Case 2: .Value = -2
 Case 3: .Value = -3
 Case 4: .Value = -4
 Case 5: .Value = -5
 Case 6: .Value = -6
 Case 7: .Value = -7
 Case 8: .Value = -8
 Case 9: .Value = -9
 Case 10: .Value = -9
 Case 11: .Value = -9
 Case 12
 For anio_x = 0 To 11
 meses(anio_x).Year = Mid(Date, 7, 10) 'meses(1).Year
 Next anio_x
 cmdmeses_Click mesDelAnio 'se queda en el mes actual
 mesesDinamicos
 End Select
 End With
End Sub

Function mesDelAnio()
 mesDelAnio = Mid(Date, 4, 2)
 mesDelAnio = mesDelAnio - 1
End Function

Private Sub mesesDinamicos()
 On Error GoTo nose:
 'tachar dias pasados
 Dim dias As Byte
 Dim ultimoDiaMes As String
 Dim anio As Integer
 For dias = 0 To 11
 meses(dias).Font.Underline = False
 meses(dias).Font.Strikethrough = False
 Next dias
 For dias = 0 To mesDelAnio - 1
 ultimoDiaMes = DateSerial(Year(Now), meses(dias).Month + 1, 0)
 ultimoDiaMes = Mid(ultimoDiaMes, 1, 2)
 meses(dias).Day = ultimoDiaMes
 meses(dias).Font.Strikethrough = True
 Next dias
 meses(mesDelAnio).Day = Day(Date)
 For dias = mesDelAnio + 1 To 11
 meses(dias).Day = 1
 meses(dias).Font.Underline = True
 Next dias
nose:
End Sub

Private Sub abc()
 Dim l_meses As Long
 Dim meses_d As Byte
 For meses_d = 1 To 26
 Dim col As New Collection
 Dim i, v, c As Integer
 l_meses = l_meses + 1
 Load cmdletra(l_meses)
 cmdletra(l_meses).Visible = True
 cmdletra(l_meses).Top = 511 * l_meses
 picabc.Height = 2280 * l_meses
 With VScroll2
 .Min = 0
 .Max = -l_meses + 3
 End With
 Next
 For i = 65 To 90
 col.Add Chr(i)
 If i = 78 Then
 col.Add "Ñ"
 End If
 Next i
 For v = 0 To 26
 c = v + 1
 cmdletra(v).Caption = col.Item(c)
 cmdletra(v).BackColor = pintarMeses()
 Next v
End Sub

Private Sub VScroll2_Change()
 picabc.Top = VScroll2.Value * 400
End Sub

Private Sub VScroll2_Scroll()
 VScroll2_Change
End Sub
