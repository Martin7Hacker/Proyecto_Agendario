VERSION 5.00
Begin VB.Form frmbienbenido 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8040
   ControlBox      =   0   'False
   Icon            =   "frmbienbenido.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   8040
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      FillColor       =   &H00800080&
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   240
      ScaleHeight     =   255
      ScaleWidth      =   7575
      TabIndex        =   5
      Top             =   1320
      Width           =   7575
      Begin VB.PictureBox pic_buf 
         BackColor       =   &H00C000C0&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   15
         TabIndex        =   6
         Top             =   0
         Width           =   15
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "10%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   3600
            TabIndex        =   7
            Top             =   15
            Width           =   615
         End
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "10%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   15
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.Image Image1 
         Height          =   720
         Left            =   6960
         Picture         =   "frmbienbenido.frx":0CCA
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Software programado por Martin Grasso Castrillo."
         ForeColor       =   &H00800080&
         Height          =   435
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   2235
      End
      Begin VB.Label Label1 
         Caption         =   "Agendario Professional v1.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   855
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label Label2 
         Caption         =   "Martinsoft"
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
         Height          =   195
         Left            =   6840
         TabIndex        =   2
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "v1.0"
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
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   1680
         Width           =   1575
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   360
      Top             =   2400
   End
End
Attribute VB_Name = "frmbienbenido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Open Source
'* System Application Software
'* Programa Bienvenido de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************
Dim contar As Integer: Const max_valor As Byte = 76

Private Sub Form_Load()
Label1.Caption = nombre_programa
End Sub

Private Sub Timer1_Timer()
contar = contar + 1
If contar = 99 Then
    contar = 0
    Unload Me
End If
cargar_descargar_buffer
End Sub

Private Sub cargar_descargar_buffer()
If pic_buf.Width > 7575 Then
   cargar = True
End If
If cargar = False Then
   pic_buf.Width = pic_buf.Width + 100
End If
If pic_buf.Width = 115 Then
   Label8.Caption = 0 & "%"
   Label9.Caption = 0 & "%"
Else
   Label8.Caption = pic_buf.Width \ _
   max_valor & "%"
   Label9.Caption = pic_buf.Width \ _
   max_valor & "%"
End If
End Sub
