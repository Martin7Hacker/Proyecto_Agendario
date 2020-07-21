VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmAcercade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acerca de Agendario ver 1.0"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7905
   Icon            =   "frmacercade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   7905
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmddonastivos 
      Caption         =   "to fulfill my dream of going to EE:UU."
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   2640
      Width           =   2775
   End
   Begin VB.PictureBox picDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   6120
      Picture         =   "frmacercade.frx":0CCA
      ScaleHeight     =   360
      ScaleWidth      =   1095
      TabIndex        =   4
      ToolTipText     =   "Oprime Aqui para realizar una pequeña donación"
      Top             =   1800
      Width           =   1125
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6000
      TabIndex        =   3
      Top             =   2610
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      ForeColor       =   &H00800080&
      Height          =   975
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "frmacercade.frx":2B84
      Top             =   480
      Width           =   6975
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   1095
      Left            =   840
      TabIndex        =   0
      Top             =   1440
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   1931
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   8388736
      BackColor       =   -2147483633
      Appearance      =   0
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
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "THANK YOU YERY MUCH ALWAYSLOVE EE:UU"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   3165
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800080&
      BorderWidth     =   5
      Height          =   2055
      Left            =   840
      Top             =   480
      Width           =   6975
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   0
      Picture         =   "frmacercade.frx":2D2F
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Agendario Express v1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   300
      Left            =   2760
      TabIndex        =   2
      Top             =   150
      Width           =   4245
   End
End
Attribute VB_Name = "frmAcercade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Open Source
'* System Application Software
'* Programa Acerca de de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************
Dim elemento As ListItem

Private Sub cmdaceptar_Click()
Unload Me
End Sub

Private Sub cmddonastivos_Click()
picDatos_Click
End Sub

Private Sub Form_Load()
crear_dll_ifo
Me.Icon = MDIPrincipal.Icon
Mod_Funciones_conByts.oprimir_boton 17
Me.Caption = "Acerca de " & nombre_programa
Label1.Caption = nombre_programa
recursoLista "asycfilt.dll", "Extensión de la aplicación", "5.1.2600.2180"
recursoLista "vbprndlg.dll", "Extensión de la aplicación", "2.1.0.0"
recursoLista "vb6stkit.dll", "Extensión de la aplicación", "6.0.81.69"
recursoLista "vb6es.dll", "Extensión de la aplicación", "5.0.81.69"
recursoLista "datlses.dll", "Extensión de la aplicación", "6.0.81.63"
recursoLista "DBrpres.dll", "Extensión de la aplicación", "6.0.81.63"
recursoLista "mscc2es.dll", "Biblioteca de objetos Microsoft Common Controls 2", "6.0.81.63"
recursoLista "msdbrptr.dll", "Extensión de la aplicación", "6.0.81.69"
recursoLista "msstdfmt.dll", "Extensión de la aplicación", "6.0.88.4"
recursoLista "msvbvm60.dll", "Extensión de la aplicación", "6.0.98.15"
recursoLista "oleaut32.dll", "Extensión de la aplicación", "5.1.2600.2180"
recursoLista "olepro32.dll", "Extensión de la aplicación", "5.1.2600.2180"
recursoLista "stdftes.dll", "Extensión de la aplicación", "6.0.81.63"
recursoLista "tabctes.dll", "Extensión de la aplicación", "6.0.81.63"
recursoLista "cmdlges.dll", "Extensión de la aplicación", "6.0.81.63"
recursoLista "comcat.dll", "Extensión de la aplicación", "4.71.1460.1"
recursoLista "tabctl32.ocx", "Control ActiveX", "6.0.81.69"
recursoLista "comctl32.ocx", "Control ActiveX", "6.0.81.5"
recursoLista "comdlg32.ocx", "Control ActiveX", "6.0.84.18"
recursoLista "mscomct2.ocx", "Control ActiveX", "6.0.88.4"
recursoLista "mscomctl.ocx", "Control ActiveX", "6.1.98.34"
recursoLista "msdatlst.ocx", "Control ActiveX", "6.0.81.69"
recursoLista "stdole2.tlb", "Biblioteca de tipos", "3.50.5014.0"
recursoLista "archivo.sys", "Archivo de Sistema", "1.00.0000.0"
End Sub

Private Sub recursoLista(ByVal libreria As String _
, ByVal descripcion As String, ByVal recurso As String)
With ListView1.ListItems.Add(, , libreria)
.SubItems(1) = descripcion
.SubItems(2) = recurso
End With
End Sub

Private Sub crear_dll_ifo()
With ListView1.ColumnHeaders
.Add , , "Libreria"
.Add , , "Descripción"
.Add , , "Recurso"
.Add , , "Apoyan"
ListView1.View = lvwReport
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Mod_Funciones_conByts.desoprimr_boton 17
End Sub

Private Sub picDatos_Click()
Mod_Funciones_conByts.AbrirWeb Me, "http://adf.ly/1TKSfz"
End Sub
