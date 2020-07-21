VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmExportaraExel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar Registros a Exel"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4575
   Icon            =   "frmExportaraExel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   4575
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmdexportar 
      Caption         =   "Exportar"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   840
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Procesador de Registros para Excel"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmExportaraExel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Open Source
'* System Application Software
'* Programa frmExportaraExel de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************

Private Sub cmdcancelar_Click()
 Unload Me
End Sub

Public Sub cmdexportar_Click()
 frmvisualizar.exportaraExel_Click
 cmdcancelar_Click
End Sub
