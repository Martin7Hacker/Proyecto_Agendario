VERSION 5.00
Begin VB.Form frmPerzonalizarDatos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perzonalizar Datos Agendario Express v1.0 "
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7170
   Icon            =   "frmPerzonalizarDatos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdboton 
      Height          =   615
      Index           =   2
      Left            =   6120
      Picture         =   "frmPerzonalizarDatos.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton cmdboton 
      Caption         =   "&Aceptar"
      Height          =   495
      Index           =   1
      Left            =   5880
      TabIndex        =   5
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdboton 
      Caption         =   "&Cancelar"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   4095
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   6975
      Begin VB.Frame Frame2 
         Height          =   3855
         Left            =   3720
         TabIndex        =   3
         Top             =   120
         Width           =   3135
         Begin VB.CommandButton cmdboton 
            Height          =   615
            Index           =   3
            Left            =   120
            Picture         =   "frmPerzonalizarDatos.frx":1994
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Restablecer "
            Top             =   3120
            Width           =   735
         End
         Begin VB.TextBox txtDato 
            BackColor       =   &H8000000B&
            Height          =   2415
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   2895
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Seleccione un dato en lista y Escrivá de debajo:"
            Height          =   375
            Left            =   120
            TabIndex        =   8
            Top             =   90
            Width           =   2895
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00800080&
            X1              =   120
            X2              =   3000
            Y1              =   480
            Y2              =   480
         End
      End
      Begin VB.ListBox List1 
         BackColor       =   &H8000000F&
         Height          =   3765
         Left            =   45
         TabIndex        =   2
         Top             =   180
         Width           =   3615
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6480
      Picture         =   "frmPerzonalizarDatos.frx":265E
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "&Perzonalizar Etiquetas de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   220
      Width           =   6975
   End
End
Attribute VB_Name = "frmPerzonalizarDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Open Source
'* System Application Software
'* Programa frmPerzonalizarDatos de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************
Private Sub restablecerDatos()
 Dim guardarDatos As String
 Open "etiquetas.opxl" For Output As 1
 Print #1, escriptar.funcion_escriptar("Name - Nombre")
 Print #1, escriptar.funcion_escriptar("Second name - Seg Nombre")
 Print #1, escriptar.funcion_escriptar("Mother's last name - Apellido M")
 Print #1, escriptar.funcion_escriptar("Last name - Apellido P")
 Print #1, escriptar.funcion_escriptar("Phone - Telefono 1")
 Print #1, escriptar.funcion_escriptar("Phone - Telefono 2")
 Print #1, escriptar.funcion_escriptar("Time to Schedule - Hora a Agendar")
 Print #1, escriptar.funcion_escriptar("Cell phone - Celular 1")
 Print #1, escriptar.funcion_escriptar("Cell phone - Celular 2")
 Print #1, escriptar.funcion_escriptar("DNI - CI")
 Print #1, escriptar.funcion_escriptar("Comments - Comentarios")
 Print #1, escriptar.funcion_escriptar("Address - Direccion")
 Print #1, escriptar.funcion_escriptar("Registration date - Fecha Registro")
 Print #1, escriptar.funcion_escriptar("Country - Pais")
 Print #1, escriptar.funcion_escriptar("Department - Departamento")
 Print #1, escriptar.funcion_escriptar("City - Ciudad")
 Print #1, escriptar.funcion_escriptar("Sex - Sexo")
 Print #1, escriptar.funcion_escriptar("Street - Calle 1")
 Print #1, escriptar.funcion_escriptar("Street - Calle 2")
 Print #1, escriptar.funcion_escriptar("Email - correo electrnico 1")
 Print #1, escriptar.funcion_escriptar("Email - correo electrnico 2")
 Print #1, escriptar.funcion_escriptar("Email - correo electrnico 3")
 Print #1, escriptar.funcion_escriptar("Age - Edad")
 Print #1, escriptar.funcion_escriptar("Birthdate - Fecha de nacimiento")
 Print #1, escriptar.funcion_escriptar("Cara libro - Facebook")
 Print #1, escriptar.funcion_escriptar("Link de Facebook 1")
 Print #1, escriptar.funcion_escriptar("Link de Facebook 2")
 Print #1, escriptar.funcion_escriptar("Twitter")
 Print #1, escriptar.funcion_escriptar("Link de Twitter 1")
 Print #1, escriptar.funcion_escriptar("Link de Twitter 2")
 Print #1, escriptar.funcion_escriptar("Number/House - Numero/Casa")
 Print #1, escriptar.funcion_escriptar("Status Civil - Estado Civil")
 Close #1
End Sub

Private Sub cmdboton_Click(Index As Integer)
 If Not (txtDato.Text = "") Then
 Select Case Index
 Case 0
 Unload Me
 Case 1
 guardarDatosLista
 Unload Me
 Case 2
 List1.List(List1.ListIndex) = txtDato.Text
 guardarDatosLista
 cmdboton(3).Enabled = True
 Case 3
 Select Case MsgBox("Quieres restablecer Datos de Agendario", _
 vbYesNo + vbInformation, "Agendario Express v1.0")
 Case (vbYes)
 restablecerDatos
 cargarDatos
 cmdboton(3).Enabled = False
 End Select
 End Select
 End If
End Sub

Private Sub cargarDatos()
 List1.Clear
 Open "etiquetas.opxl" For Input As 1
 Dim etiqueta As String
 Do While Not EOF(1)
 Line Input #1, etiqueta
 List1.AddItem escriptar.funcion_desescriptar(etiqueta)
 Loop
 Close #1
End Sub

Private Sub guardarDatosLista()
 Dim guardarDatos As Integer
 Open "etiquetas.opxl" For Output As 1
 For guardarDatos = 0 To List1.ListCount - 1
 Print #1, escriptar.funcion_escriptar(List1.List(guardarDatos))
 Next
 Close #1
End Sub

Private Sub Form_Load()
 cargarDatos
End Sub

Private Sub List1_Click()
 txtDato.Text = List1.List(List1.ListIndex)
End Sub
