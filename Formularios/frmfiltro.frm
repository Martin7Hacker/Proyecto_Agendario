VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmfiltro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtrar Columnas"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7860
   Icon            =   "frmfiltro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmfiltro.frx":0CCA
   ScaleHeight     =   4260
   ScaleWidth      =   7860
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdninguno 
      Caption         =   "&Limpiar"
      Height          =   375
      Left            =   3960
      TabIndex        =   7
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton cmdvertodo 
      Caption         =   "&Ver todos los registros"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   60
      Left            =   120
      ScaleHeight     =   0
      ScaleWidth      =   7635
      TabIndex        =   2
      Top             =   360
      Width           =   7695
   End
   Begin VB.ListBox List1 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3300
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   480
      Width           =   2535
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3300
      Left            =   2760
      TabIndex        =   3
      Top             =   480
      Width           =   5085
      _ExtentX        =   8969
      _ExtentY        =   5821
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   8388736
      BorderStyle     =   1
      Appearance      =   1
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
      Picture         =   "frmfiltro.frx":0D10
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7200
      Picture         =   "frmfiltro.frx":601C2
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "--Simulación de Columnas de Datos--"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3960
      TabIndex        =   8
      Top             =   120
      Width           =   2595
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Columnas a &Ver"
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
End
Attribute VB_Name = "frmfiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Open Source
'* System Application Software
'* Programa frmfiltro de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************
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

Private Sub cargar_datos()
 With List1
 List1.Clear
 Open "etiquetas.opxl" For Input As 1
 Dim etiqueta As String
 Do While Not EOF(1)
 Line Input #1, etiqueta
 .AddItem escriptar.funcion_desescriptar(etiqueta)
 Loop
 Close #1
 End With
End Sub

Private Sub cargarDatos()
 Dim r As Long
 ListView1.ListItems.Clear
 Dim i As Long
 For i = 1 To agenda.nombre.Count
 If Not (agenda.nombre.Count = 1) Then
 End If
 Set elemento = ListView1.ListItems.Add(, , agenda.nombre.Item(i).Key)
 frmcargando.ProgressBar1.Value = i
 r = i
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
End Sub

Private Sub cmdaplicar_Click()
 aplicarControl
End Sub

Private Sub cmdcancelar_Click()
 Unload Me
End Sub

Private Sub cmdninguno_Click()
 verOno False
 irPrimerRegistro
End Sub

Private Sub cmdvertodo_Click()
 verOno True
 irPrimerRegistro
End Sub

Private Sub Form_Load()
 Me.Icon = MDIPrincipal.Icon
 agregar_registros
 cargarDatos
 cargar_datos
 verOno True
 List1_Click
 frmfiltro.ListView1.Picture = LoadPicture("img\simulador.bmp")
End Sub

Private Sub aplicarControl()
 Dim y, x As Integer
 For y = 1 To 32
 x = y - 1
 lista frmvisualizar.ListView1, x, y
 frmvisualizar.visualizar_todos_los_registros
 Next
End Sub

Private Sub List1_Click()
 Dim y As Integer
 Dim x As Integer
 For y = 1 To 32
 x = y - 1
 lista ListView1, x, y
 Next
End Sub

Public Sub lista(ByVal control As ListView, ByVal lis _
As Integer, ByVal con As Integer)
 If List1.Selected(lis) = True Then
 control.ColumnHeaders(con).Width = 1500
 ElseIf List1.Selected(lis) = False Then
 control.ColumnHeaders(con).Width = 0
 End If
End Sub

Private Sub verOno(ByVal ver As Boolean)
 On Error GoTo nose
 Dim x As Integer
 For x = 0 To 31
 List1.Selected(x) = ver
 Next x
nose:
End Sub

Private Sub irPrimerRegistro()
 List1.ListIndex = 0
End Sub
