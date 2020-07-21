VERSION 5.00
Begin VB.Form frmconsultas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Explorar Martinsoft"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7065
   Icon            =   "frmconsultas.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   7065
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmddonar 
      Height          =   495
      Left            =   3000
      Picture         =   "frmconsultas.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   3360
      Width           =   1095
   End
   Begin VB.OptionButton optionx 
      Caption         =   "Mis Programas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   25
      Left            =   3720
      TabIndex        =   27
      Top             =   2760
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "Mi Book"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   24
      Left            =   600
      TabIndex        =   26
      Top             =   3000
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "Wikiwhatsapp"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   23
      Left            =   3720
      TabIndex        =   25
      Top             =   2520
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "Captcha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   22
      Left            =   600
      TabIndex        =   24
      Top             =   2760
      Width           =   3015
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdexportar 
      Caption         =   "&Explorar "
      Height          =   495
      Left            =   5640
      TabIndex        =   22
      Top             =   3360
      Width           =   1215
   End
   Begin VB.OptionButton optionx 
      Caption         =   "Acerca de Martinsoft"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   21
      Left            =   3720
      TabIndex        =   21
      Top             =   3000
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "&Música"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   20
      Left            =   3720
      TabIndex        =   20
      Top             =   2280
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "&Libreria de Software"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   19
      Left            =   3720
      TabIndex        =   19
      Top             =   2040
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "&Mis Estudios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   18
      Left            =   3720
      TabIndex        =   18
      Top             =   1800
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "&Mis Inventos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   17
      Left            =   3720
      TabIndex        =   17
      Top             =   1560
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "&Calculador de Resistencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   16
      Left            =   3720
      TabIndex        =   16
      Top             =   1320
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "Juego TATETI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   15
      Left            =   3720
      TabIndex        =   15
      Top             =   1080
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "Generador de Tablas de Múltiplicar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   14
      Left            =   3720
      TabIndex        =   14
      Top             =   840
      Width           =   3375
   End
   Begin VB.OptionButton optionx 
      Caption         =   "Mi Galeria Fotografica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   13
      Left            =   3720
      TabIndex        =   13
      Top             =   600
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "My YouTube"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   12
      Left            =   3720
      TabIndex        =   12
      Top             =   360
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "&Curso de Agendario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   11
      Left            =   3720
      TabIndex        =   11
      Top             =   120
      Width           =   3135
   End
   Begin VB.OptionButton optionx 
      Caption         =   "&Descargas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   10
      Left            =   600
      TabIndex        =   10
      Top             =   2520
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "&Donar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   9
      Left            =   600
      TabIndex        =   9
      Top             =   2280
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "&Cursos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   8
      Left            =   600
      TabIndex        =   8
      Top             =   2040
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "&App"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   7
      Left            =   600
      TabIndex        =   7
      Top             =   1800
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "&Códigos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   6
      Left            =   600
      TabIndex        =   6
      Top             =   1560
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "&Clima"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   5
      Top             =   1320
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "Vacaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   4
      Left            =   600
      TabIndex        =   4
      Top             =   1080
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "Electrónica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   3
      Left            =   600
      TabIndex        =   3
      Top             =   840
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "Mensajeria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "Playlist"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "donativos por:"
      ForeColor       =   &H00400040&
      Height          =   255
      Left            =   1680
      TabIndex        =   29
      Top             =   3480
      Width           =   2535
   End
End
Attribute VB_Name = "frmconsultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Open Source
'* System Application Software
'* Programa Consultas de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************
Private Sub cmdcancelar_Click()
Unload Me
End Sub

Private Sub cmddonar_Click()
Mod_Funciones_conByts.AbrirWeb Me, "http://adf.ly/1TJlyy"
End Sub

Private Sub cmdexportar_Click()
On Error GoTo nose
Dim op As Byte
 For op = 0 To 25
 If optionx(op).Value = True Then
    Select Case (op)
           Case 0 'inicio
            Mod_Funciones_conByts.AbrirWeb Me, _
            "http://adf.ly/1TJk7K"
           Case 1 'Playlist
            Mod_Funciones_conByts.AbrirWeb Me, _
            "http://adf.ly/1TJkz7"
           Case 2 'Mensajeria
            Mod_Funciones_conByts.AbrirWeb Me, _
            "http://adf.ly/1TJl76"
           Case 3 'Electrónica
            Mod_Funciones_conByts.AbrirWeb Me, _
            "http://adf.ly/1TJlDh"
           Case 4 'Vacaciónes
            Mod_Funciones_conByts.AbrirWeb Me, _
            "http://adf.ly/1TJlK9"
           Case 5 'Clima
            Mod_Funciones_conByts.AbrirWeb Me, _
            "http://adf.ly/1TJlXi"
           Case 6 'Códigos
            Mod_Funciones_conByts.AbrirWeb Me, _
            "http://adf.ly/1TJlds"
           Case 7 'App
            Mod_Funciones_conByts.AbrirWeb Me, _
            "http://adf.ly/1TJliq"
           Case 8 'Cursos
            Mod_Funciones_conByts.AbrirWeb Me, _
            "http://adf.ly/1TJls6"
           Case 9 'Donar
            Mod_Funciones_conByts.AbrirWeb Me, _
            "http://adf.ly/1TJlyy"
           Case 10 'Descargas
            Mod_Funciones_conByts.AbrirWeb Me, _
            "http://adf.ly/1TJmJn"
           Case 11 'Cursos de Agendario
            Mod_Funciones_conByts.AbrirWeb Me, _
            "http://adf.ly/1TJmhG"
           Case 12 'My YouTube
            frmMisCanales.Show 1
           Case 13 'Mi Galeria Fotografica
            Mod_Funciones_conByts.AbrirWeb Me, _
            "http://adf.ly/1TJnvV"
           Case 14 'Generador de Tablas de Múltiplicar
            Mod_Funciones_conByts.AbrirWeb Me, _
            "http://adf.ly/1TJof5"
           Case 15 'Juego TATETI
            Mod_Funciones_conByts.AbrirWeb Me, _
            "http://adf.ly/1TJqT0"
           Case 16 'Calculador de Resistencia
            Mod_Funciones_conByts.AbrirWeb Me, _
            "http://adf.ly/1TJpWV"
           Case 17 'Mis Inventos
            Mod_Funciones_conByts.AbrirWeb Me, _
            "http://adf.ly/1TJqlO"
           Case 18 'Mis Estudios
            Mod_Funciones_conByts.AbrirWeb Me, _
            "http://adf.ly/1TJr18"
           Case 19 'Libreria de Software
            Mod_Funciones_conByts.AbrirWeb Me, _
            "http://adf.ly/1TJrOF"
           Case 20 'Música
            Mod_Funciones_conByts.AbrirWeb Me, _
            "http://adf.ly/1TJrZI"
           Case 21 'Acerca de Martinsoft
            Mod_Funciones_conByts.AbrirWeb Me, _
            "http://adf.ly/1TJrkA"
           Case 22 'Captcha
            Mod_Funciones_conByts.AbrirWeb Me, _
            "http://adf.ly/1TJsEJ"
           Case 23 'Wikiwhatsapp
            Mod_Funciones_conByts.AbrirWeb Me, _
            "http://adf.ly/1TJtE4"
           Case 24 'Mi Book
            Mod_Funciones_conByts.AbrirWeb Me, _
            "http://adf.ly/1TJtoh"
           Case 25 'Mis Programas
            Mod_Funciones_conByts.AbrirWeb Me, _
            "http://adf.ly/1TJuFe"
    End Select
 End If
 Next op
nose:
End Sub

Private Sub Label1_Click()
cmddonar_Click
End Sub


