VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmcopiaSeguridad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear copia de seguridad de base de datos "
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7485
   Icon            =   "frmcopiaSeguridad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   7485
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog dialogoGuardar 
      Left            =   2640
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAlmacenamiento 
      Caption         =   "Búscar destino"
      Height          =   255
      Left            =   5760
      TabIndex        =   7
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   5760
      TabIndex        =   6
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton cmdguardar 
      Caption         =   "&Guardar"
      Enabled         =   0   'False
      Height          =   255
      Left            =   5760
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtdestino 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Text            =   "C:\"
      Top             =   450
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.PictureBox Picture1 
         Height          =   735
         Left            =   5400
         ScaleHeight     =   675
         ScaleWidth      =   0
         TabIndex        =   3
         Top             =   120
         Width           =   60
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmcopiaSeguridad.frx":0CCA
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Dirección de Almacenamiento de la BD:"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   2820
      End
   End
End
Attribute VB_Name = "frmcopiaSeguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Open Source
'* System Application Software
'* Programa frmcopiaseguridad de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************
Private Sub cmdaceptar_Click()
Unload Me
End Sub

Private Sub cmdAlmacenamiento_Click()
 With dialogoGuardar
 If .CancelError = False Then
 .DialogTitle = "Guardar"
 .Filter = "Formato de Base de Datos de Oplx(*.Oplx)|*.Oplx"
 .ShowSave
 If .FileName = "" Then
 cmdguardar.Enabled = False
 MsgBox "Escrive un nombre de Archivo para Guardar", vbInformation
 End If
 If .FileName <> "" Then
 txtdestino.Text = .FileName
 .FileName = ""
 cmdguardar.Enabled = True
 End If
 End If
End With
End Sub

Private Sub cmdcancelar_Click()
Unload Me
End Sub

Private Sub cmdguardar_Click()
On Error GoTo nose
    Call FileCopy(App.Path & "\bd.opxl", txtdestino.Text)
    MsgBox " El Archivo de BD Opxl se ha copiado en: " & _
    vbCrLf & txtdestino.Text, vbInformation
nose:
End Sub

Private Sub txtdestino_Change()
 On Error GoTo nose
 If txtdestino.Text = "" Then
 txtdestino.Enabled = False
 Else
 txtdestino.Enabled = True
 End If
nose:
End Sub
