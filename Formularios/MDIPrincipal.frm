VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H00404040&
   Caption         =   "Agendario Express v1.0"
   ClientHeight    =   8295
   ClientLeft      =   120
   ClientTop       =   750
   ClientWidth     =   12945
   Icon            =   "MDIPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIPrincipal.frx":0CCA
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12945
      _ExtentX        =   22834
      _ExtentY        =   1535
      ButtonWidth     =   2037
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   19
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   ""
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Agendar[F1]"
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "&Agendar - Schedule[F1]"
            Object.Tag             =   ""
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Búscar[F2]"
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "&Búscar - Search[F2]"
            Object.Tag             =   ""
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Modificar[F3]"
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "&Modificar - Modify[F3] "
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Eliminar[F4]"
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Eliminar - Remove[F4]"
            Object.Tag             =   ""
            ImageIndex      =   4
            Style           =   1
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&P. Virtual[F5]"
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Planilla Virtual - Virtual Sheet[F5]"
            Object.Tag             =   ""
            ImageIndex      =   5
            Style           =   1
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Exportar[F6]"
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "&Exportar - To export[F6]"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Imprimir[F7]"
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "&Imprimir - Print[F7]"
            Object.Tag             =   ""
            ImageIndex      =   7
            Style           =   1
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Salir[F8]"
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Salir - Exit[F8]"
            Object.Tag             =   ""
            ImageIndex      =   8
            Style           =   1
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Acerca[F9]"
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Acerca de - About[F9] "
            Object.Tag             =   ""
            ImageIndex      =   9
            Style           =   1
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Donar: $[F11]"
            Key             =   ""
            Description     =   ""
            Object.ToolTipText     =   "Donar - Donate[F11]"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar StatusBar3 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   1
      Top             =   7890
      Width           =   12945
      _ExtentX        =   22834
      _ExtentY        =   714
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   4287
            MinWidth        =   4287
            Picture         =   "MDIPrincipal.frx":25A84C
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Text            =   "Númerico"
            TextSave        =   "Númerico"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Enabled         =   0   'False
            Text            =   "Mayúsculas"
            TextSave        =   "Mayúsculas"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            TextSave        =   "27/07/2016"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            TextSave        =   "16:05"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6240
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   10
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrincipal.frx":25D9C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrincipal.frx":25E6A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrincipal.frx":25F37A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrincipal.frx":260054
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrincipal.frx":260D2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrincipal.frx":261A08
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrincipal.frx":2626E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrincipal.frx":2633BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrincipal.frx":264096
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIPrincipal.frx":264D70
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu Elimnar 
      Caption         =   "Elimnar - Remove"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu eleimniartodo 
         Caption         =   "&Elimniar Todo - Delete all "
      End
      Begin VB.Menu SoloSelecionado 
         Caption         =   "&Sólo Seleccionado - only Selected"
      End
   End
   Begin VB.Menu archivo 
      Caption         =   "&Archivo - File"
      Begin VB.Menu Agendar 
         Caption         =   "&Agendar - Schedule"
         Shortcut        =   {F1}
      End
      Begin VB.Menu buscar 
         Caption         =   "&Búscar - Search"
         Shortcut        =   {F2}
      End
      Begin VB.Menu Modificar 
         Caption         =   "&Modificar - Modify"
         Shortcut        =   {F3}
      End
      Begin VB.Menu Elimniar 
         Caption         =   "&Eliminar - Remove"
         Shortcut        =   {F4}
      End
      Begin VB.Menu planillavirtual 
         Caption         =   "&Planilla Virtual - Virtual Sheet"
         Shortcut        =   {F5}
      End
      Begin VB.Menu Exportarx 
         Caption         =   "&Exportar - To export"
         Shortcut        =   {F6}
      End
      Begin VB.Menu imprimir 
         Caption         =   "&Imprimir - Print"
         Shortcut        =   {F7}
      End
      Begin VB.Menu DonaralProyecto 
         Caption         =   "&Donar al Proyecto - Donate to Project"
         Shortcut        =   {F11}
      End
      Begin VB.Menu salir 
         Caption         =   "&Salir - Exit"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu explorarmartinsoft 
      Caption         =   "&Explorar Martinsoft - explore Martinsoft"
   End
   Begin VB.Menu clima 
      Caption         =   "&Clima - Weather"
   End
   Begin VB.Menu seguridad 
      Caption         =   "&Seguridad - Security"
   End
   Begin VB.Menu ayuda 
      Caption         =   "&Ayuda - Help"
      Begin VB.Menu donaci 
         Caption         =   "&Donación $ para Software Agendario Express v1.0 - Donation $ Agendario Express Software v1.0"
      End
      Begin VB.Menu ayudadeagendario 
         Caption         =   "&Ayuda de Agendario - Help Agendario"
         Shortcut        =   {F12}
      End
      Begin VB.Menu AcercadeAgendariox 
         Caption         =   "&Acerca de Agendario - About Agendario Express v1.0"
         Shortcut        =   {F9}
      End
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Open Source
'* System Application Software
'* Programa MDIPrincipal de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************

Private Sub acercadeagendario_Click()
 frmAcercade.Show 1
End Sub

Private Sub AcercadeAgendariox_Click()
 frmAcercade.Show 1
End Sub

Private Sub Agendar_Click()
 Me.crearNuevaHoja
End Sub

Private Sub ayudadeagendario_Click()
 Mod_Funciones_conByts.openAyuda Me
End Sub

Private Sub buscar_Click()
 Me.buscarEnBD
End Sub

Private Sub clima_Click()
 Mod_Funciones_conByts.consultartiempo Me
End Sub

Private Sub donaci_Click()
 Mod_Funciones_conByts.oprimir_boton 19
 Select Case MsgBox("Quieres Realizar una Donacion para Ayudar al Proyecto Agendario Express v1.0", vbYesNo + vbInformation, nombre_programa)
 Case (vbYes)
 Mod_Funciones_conByts.desoprimr_boton 19
 Mod_Funciones_conByts.AbrirWeb Me, "http://adf.ly/1TJlyy"
 Case (vbNo)
 Mod_Funciones_conByts.desoprimr_boton 19
 End Select
End Sub

Private Sub DonaralProyecto_Click()
 donaci_Click
End Sub

Private Sub Elimniar_Click()
 Me.elimnarSelecionado
End Sub

Private Sub explorarmartinsoft_Click()
 frmconsultas.Show 1
End Sub

Private Sub Exportarx_Click()
 Me.ExportarArchivo
End Sub

Private Sub imprimir_Click()
 Me.Inprimir
End Sub

Private Sub MDIForm_Load()
 On Error GoTo nose
 ModDatos.abrir_archivo
 ModPrincipal.Abrir
 frmbienbenido.Show 1
 MDIPrincipal.Picture = LoadPicture("img\agendario.bmp")
 If y_seguridad.y_poderver = "" Then
 funcion_mostrar_formulario frmvisualizar, Button, 0
 Else
 If Not (frmvisualizar.ListView1.Enabled = False) Then
 frmpoderver.Show 1
 End If
 End If
nose:
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
 On Error GoTo nose
 Cancel = 1
 Select Case MsgBox("Quieres salir de " & nombre_programa, _
 vbInformation + vbYesNo, nombre_programa)
 Case (vbYes)
 End
 End Select
nose:
End Sub

Private Sub Modificar_Click()
 frmvisualizar.cmdModificar_Click
End Sub

Private Sub planillavirtual_Click()
 Me.MostrarPlanillaVirtual
End Sub

Private Sub salir_Click()
 Unload Me
End Sub

Private Sub seguridad_Click()
 If y_seguridad.y_iniciodel = "" Then
 frmseguridad.Show 1
 Else
 frmseguridadx.Show 1
 End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
 Select Case Button.Index
 Case (2)
 If y_seguridad.y_creacion = "" Then
 funcion_mostrar_formulario frmdatos, Button, 0
 Else
 frmcrearx.Show 1
 End If
 Case (4)
 If y_seguridad.y_busqueda = "" Then
 If frmvisualizar.Visible = False Then
 funcion_mostrar_formulario frmvisualizar, Button, 0
 funcion_mostrar_formulario frmbusqueda, Button, 0
 Else
 funcion_mostrar_formulario frmbusqueda, Button, 0
 End If
 Else
 frmpoderbuscar.Show 1
 End If
 Case (6)
 frmvisualizar.cmdModificar_Click
 Case (8)
 If y_seguridad.y_eliminarSeleccionado = "" Then
 frmvisualizar.eliminarselecionado_Click
 Else
 frmElimnarSelecionado.Show 1
 End If
 Case (10)
 funcion_mostrar_formulario frmvisualizar, Button, 0
 Case (12)
 If y_seguridad.y_poderExportar = "" Then
 funcion_mostrar_formulario frmexportar_datos, Button, 1
 frmexportar_datos.Show 1
 Else
 frmExportarDatos.Show 1
 End If
 Case (13)
 If y_seguridad.y_poderImprimir = "" Then
 frmImprimir.Show 1
 Else
 frmpoderImprimir.Show 1
 End If
 Case (15)
 Select Case MsgBox("Estas Seguro de Que Quieres Salir de" & nombre_programa & "?", vbInformation + vbYesNo, nombre_programa)
 Case (vbYes)
 Mod_Funciones_conByts.desoprimr_boton 15
 End
 Case (vbNo)
 Mod_Funciones_conByts.desoprimr_boton 15
 End Select
 Case (17)
 funcion_mostrar_formulario frmAcercade, Button, 1
 Case (20)
 frmFreeSoftware.Show 1
 Case (19)
 donaci_Click
 End Select
End Sub

Private Sub funcion_mostrar_formulario(ByVal formulario As Form, _
ByVal Button As ComctlLib.Button, ByVal e As Byte)
 Select Case Button.Value
 Case (tbrPressed)
 Select Case e
 Case (0)
 formulario.Show
 Case (1)
 formulario.Show 1
 End Select
 Case (tbrUnpressed)
 Unload formulario
 End Select
End Sub

Public Sub crearNuevaHoja()
 If y_seguridad.y_creacion = "" Then
 Mod_Funciones_conByts.oprimir_boton 2
 frmdatos.Show 1
 Else
 frmcrearx.Show 1
 End If
End Sub

Public Sub buscarEnBD()
 If y_seguridad.y_busqueda = "" Then
 If frmvisualizar.Visible = False Then
 Mod_Funciones_conByts.oprimir_boton 4
 frmvisualizar.Show
 frmbusqueda.Show 1
 Else
 Mod_Funciones_conByts.oprimir_boton 4
 frmbusqueda.Show 1
 End If
 Else
 frmpoderbuscar.Show 1
 End If
End Sub

Public Sub elimnarSelecionado()
 If y_seguridad.y_eliminarSeleccionado = "" Then
 frmvisualizar.eliminarselecionado_Click
 Else
 frmElimnarSelecionado.Show 1
 End If
End Sub

Public Sub MostrarPlanillaVirtual()
frmvisualizar.Show
End Sub

Public Sub ExportarArchivo()
 If y_seguridad.y_poderExportar = "" Then
 Mod_Funciones_conByts.oprimir_boton 12
 frmexportar_datos.Show 1
 Else
 frmExportarDatos.Show 1
 End If
End Sub

Public Sub Inprimir()
 If y_seguridad.y_poderImprimir = "" Then
 Mod_Funciones_conByts.oprimir_boton 13
 frmImprimir.Show 1
 Else
 frmpoderImprimir.Show 1
 End If
End Sub

