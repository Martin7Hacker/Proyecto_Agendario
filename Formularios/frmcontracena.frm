VERSION 5.00
Begin VB.Form frmcontracena 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingresar Contraceña"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3855
   Icon            =   "frmcontracena.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   3855
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   960
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   360
      Width           =   3615
   End
   Begin VB.Label Label2 
      Caption         =   "Contraceña de Confirmación:"
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Contraceña Base:"
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1275
   End
End
Attribute VB_Name = "frmcontracena"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Open Source
'* System Application Software
'* Programa Contraceña de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************
Public control_fx As Integer

Private Sub cmdaceptar_Click()
 If Text1.Text = Text2.Text Then
 MsgBox "La contraceña es ha confirmado con exito", vbInformation
 frmseguridad.controlPasarContracena Text1.Text, control_fx
 Text1.Text = frmcontracena.Text1.Text
 If frmseguridad.Option1(0).Value = True Then
    frmseguridad.Text1.Text = Text1.Text
    frmseguridad.Text2.Text = Text1.Text
    frmseguridad.Text3.Text = Text1.Text
    frmseguridad.Text4.Text = Text1.Text
    frmseguridad.Text5.Text = Text1.Text
    frmseguridad.Text6.Text = Text1.Text
    frmseguridad.Text7.Text = Text1.Text
    frmseguridad.Text8.Text = Text1.Text
    frmseguridad.Text9.Text = Text1.Text
    frmseguridad.Text10.Text = Text1.Text
    frmseguridad.Text11.Text = Text1.Text
 End If
   Unload Me
   Else
   MsgBox "La contraceña base y la contraceña de confirmación son diferentes ingrese una igual en ambos casos", vbCritical
 End If
End Sub

