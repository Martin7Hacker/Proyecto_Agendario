VERSION 5.00
Begin VB.Form frmMisCanales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mis Canales de YouTube"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3795
   Icon            =   "frmMisCanales.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   3795
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdexportar 
      Caption         =   "&Explorar "
      Height          =   495
      Left            =   2520
      TabIndex        =   11
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   2640
      Width           =   1215
   End
   Begin VB.OptionButton optionx 
      Caption         =   "Google+ Martin Grasso"
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
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "Google+ Solo Electrónica .TV"
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
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "Google+ Sintaxisxd"
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
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "Google+ Lorest"
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
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "Google+ The Hackers"
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
      Left            =   120
      TabIndex        =   5
      Top             =   2280
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "Martin Grasso"
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
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "Solo Electrónica .TV"
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
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "Sintaxisxd"
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
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "Lorest"
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
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3015
   End
   Begin VB.OptionButton optionx 
      Caption         =   "The Hackers"
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
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3015
   End
End
Attribute VB_Name = "frmMisCanales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Open Source
'* System Application Software
'* Programa frmMisCanales de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************

Private Sub cmdcancelar_Click()
 Unload Me
End Sub

Private Sub cmdexportar_Click()
 On Error GoTo nose
 Dim op As Byte
 For op = 0 To 9
 If optionx(op).Value = True Then
 Select Case (op)
 Case 0 'Martin Grasso
 Mod_Funciones_conByts.AbrirWeb Me, _
 "http://adf.ly/1TK2rZ"
 Case 1 'Solo Electrónica .TV
 Mod_Funciones_conByts.AbrirWeb Me, _
 "http://adf.ly/1MEg11"
 Case 2 'Sintaxisxd
 Mod_Funciones_conByts.AbrirWeb Me, _
 "http://adf.ly/1TK3Ot"
 Case 3 'Lorest
 Mod_Funciones_conByts.AbrirWeb Me, _
 "http://adf.ly/1TK4FQ"
 Case 4 'The Hackers
 Mod_Funciones_conByts.AbrirWeb Me, _
 "http://adf.ly/1TK4b4"
 Case 5 'Google+ The Hackers
 Mod_Funciones_conByts.AbrirWeb Me, _
 "http://adf.ly/1TK5JJ"
 Case 6 'Google+ Lorest
 Mod_Funciones_conByts.AbrirWeb Me, _
 "http://adf.ly/1TK5Uf"
 Case 7 'Google+ Sintaxisxd
 Mod_Funciones_conByts.AbrirWeb Me, _
 "http://adf.ly/1TK5dG"
 Case 8 'Google+ Solo Electrónica .TV
 Mod_Funciones_conByts.AbrirWeb Me, _
 "http://adf.ly/1TK5lz"
 Case 9 'Google+ Martin Grasso
 Mod_Funciones_conByts.AbrirWeb Me, _
 "http://adf.ly/1TK6PP"
 End Select
 End If
 Next op
nose:
End Sub
