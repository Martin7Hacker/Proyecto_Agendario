VERSION 5.00
Begin VB.Form cmdingresar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Propiedades"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3615
   Icon            =   "cmdingresar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   3615
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   2175
      TabIndex        =   12
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton Command7 
      Caption         =   "colorear"
      Height          =   255
      Left            =   2280
      TabIndex        =   11
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Aplicar"
      Height          =   375
      Left            =   2640
      TabIndex        =   10
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Borde de  Tabla:    SI"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   3375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Linea Secundaria: SI"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   3375
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2310
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1000
      Width           =   1125
   End
   Begin VB.CommandButton Command2 
      Caption         =   "colorear"
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800080&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   2175
      TabIndex        =   3
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Linea Primaria:       SI"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Label Label2 
      Caption         =   "Tamaño de Letra:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1050
      Width           =   3330
   End
   Begin VB.Label Label1 
      Caption         =   "Texto:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3330
   End
End
Attribute VB_Name = "cmdingresar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Open Source
'* System Application Software
'* Programa Ingresar de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************
Private Sub agregar_elementosenlista()
 With Combo1
 .AddItem "h1"
 .AddItem "h2"
 .AddItem "h3"
 .AddItem "h4"
 .AddItem "h5"
 .AddItem "h6"
 End With
End Sub

Private Sub Form_Load()
Combo1.Clear
agregar_elementosenlista
End Sub
