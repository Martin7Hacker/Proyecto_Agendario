VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmexportar_datos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar Datos."
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7230
   Icon            =   "frmexportar_datos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   7230
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdperzonalizar 
      Caption         =   "+"
      Height          =   315
      Left            =   3360
      TabIndex        =   22
      Top             =   2199
      Visible         =   0   'False
      Width           =   740
   End
   Begin MSComDlg.CommonDialog dialogoGuardar 
      Left            =   3360
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture6 
      Height          =   2565
      Left            =   120
      ScaleHeight     =   2505
      ScaleWidth      =   0
      TabIndex        =   19
      Top             =   480
      Width           =   60
   End
   Begin VB.PictureBox Picture5 
      Height          =   2565
      Left            =   7080
      ScaleHeight     =   2505
      ScaleWidth      =   0
      TabIndex        =   18
      Top             =   480
      Width           =   60
   End
   Begin VB.CommandButton cmdGuardarHtml 
      Caption         =   "&Guardar en..."
      Height          =   405
      Left            =   4320
      TabIndex        =   17
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Guardar en..."
      Height          =   405
      Left            =   4320
      TabIndex        =   16
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&Guardar en..."
      Height          =   405
      Left            =   4320
      TabIndex        =   15
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton cmdExportarHTML 
      Caption         =   "&Exportar..."
      Enabled         =   0   'False
      Height          =   405
      Left            =   5760
      TabIndex        =   14
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "&Exportar..."
      Enabled         =   0   'False
      Height          =   405
      Left            =   5760
      TabIndex        =   13
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6000
      TabIndex        =   11
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   720
      TabIndex        =   9
      Text            =   "C:\"
      Top             =   2520
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      Height          =   60
      Index           =   4
      Left            =   120
      ScaleHeight     =   0
      ScaleWidth      =   6960
      TabIndex        =   7
      Top             =   3000
      Width           =   7020
   End
   Begin VB.PictureBox Picture1 
      Height          =   60
      Index           =   2
      Left            =   120
      ScaleHeight     =   0
      ScaleWidth      =   6960
      TabIndex        =   6
      Top             =   2160
      Width           =   7020
   End
   Begin VB.CommandButton cmdexportarexel 
      Caption         =   "&Exportar..."
      Enabled         =   0   'False
      Height          =   405
      Left            =   5760
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Text            =   "C:\"
      Top             =   1680
      Width           =   3375
   End
   Begin VB.PictureBox Picture1 
      Height          =   60
      Index           =   1
      Left            =   120
      ScaleHeight     =   0
      ScaleWidth      =   6960
      TabIndex        =   2
      Top             =   480
      Width           =   7020
      Begin VB.PictureBox Picture2 
         Height          =   1695
         Left            =   -1080
         ScaleHeight     =   1635
         ScaleWidth      =   4875
         TabIndex        =   12
         Top             =   -1680
         Width           =   4935
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   60
      Index           =   0
      Left            =   120
      ScaleHeight     =   0
      ScaleWidth      =   6960
      TabIndex        =   1
      Top             =   1320
      Width           =   7020
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   720
      TabIndex        =   0
      Text            =   "C:\"
      Top             =   840
      Width           =   3375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Exportar Datos a formato de Texto :"
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
      Left            =   720
      TabIndex        =   21
      Top             =   600
      Width           =   3375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800080&
      BorderWidth     =   2
      X1              =   120
      X2              =   7080
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label4 
      Caption         =   "Sistema de exportación de Archivos de Agendario"
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
      TabIndex        =   20
      Top             =   120
      Width           =   7005
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   200
      Picture         =   "frmexportar_datos.frx":0CCA
      Top             =   2320
      Width           =   480
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Exportar datos a código Html:"
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
      Index           =   1
      Left            =   720
      TabIndex        =   8
      Top             =   2280
      Width           =   3495
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   195
      Picture         =   "frmexportar_datos.frx":1994
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Exportar Datos a formato Excel :"
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
      Left            =   720
      TabIndex        =   3
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   200
      Picture         =   "frmexportar_datos.frx":265E
      Top             =   705
      Width           =   480
   End
End
Attribute VB_Name = "frmexportar_datos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Open Source
'* System Application Software
'* Programa frmexportar de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************
Private Sub cmdcancelar_Click()
 Unload Me
End Sub

Private Sub cmdexportar_Click()
 Call Exportar_ListView(frmvisualizar.ListView1, Text1(0).Text, vbTab)
End Sub

Private Sub cmdexportarexel_Click()
 frmExportaraExel.cmdexportar_Click
End Sub

Private Sub cmdExportarHTML_Click()
 Call ModExportaraHTML.ExportarHTML_Chrome(Text1(1).Text, frmvisualizar.ListView1, _
 "Archivos de Agendario v1.0", "Martinsoft Software")
End Sub

Private Sub cmdGuardarHtml_Click()
 With dialogoGuardar
 If .CancelError = False Then
 .DialogTitle = "Guardar"
 .Filter = "Formato de código HTML (*.html)|*.html"
 .ShowSave
 If .FileName = "" Then
 cmdExportarHTML.Enabled = False
 MsgBox "Escrive un nombre de Archivo para Guardar", vbInformation
 End If
 If .FileName <> "" Then
 Text1(1).Text = .FileName
 .FileName = ""
 cmdExportarHTML.Enabled = True
 End If
 End If
 End With
End Sub

Private Sub cmdperzonalizar_Click()
frmperzonalizar.Show 1
End Sub

Private Sub Command6_Click()
 With dialogoGuardar
 If .CancelError = False Then
 .DialogTitle = "Guardar"
 .Filter = "Formato de Texto (*.txt)|*.txt|todos los Archivos (*.*)|*.*|"
 .ShowSave
 If .FileName = "" Then
 MsgBox "Escrive un nombre de Archivo para Guardar", vbInformation
 End If
 If .FileName <> "" Then
 Text1(0).Text = .FileName
 .FileName = ""
 cmdExportar.Enabled = True
 End If
 End If
 End With
End Sub

Private Sub Command7_Click()
 With dialogoGuardar
 If .CancelError = False Then
 .DialogTitle = "Guardar"
 .Filter = "Formato de Libros de Microsoft Exel(*.xls)|*.xls"
 .ShowSave
 If .FileName = "" Then
 cmdexportarexel.Enabled = False
 MsgBox "Escrive un nombre de Archivo para Guardar", vbInformation
 End If
 If .FileName <> "" Then
 Text2.Text = .FileName
 .FileName = ""
 Exportar_ListViewExel
 cmdexportarexel.Enabled = True
 End If
 End If
 End With
End Sub

Private Sub Form_Load()
 Mod_Funciones_conByts.oprimir_boton 12
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Mod_Funciones_conByts.desoprimr_boton 12
End Sub

Public Sub Exportar_ListView(ListView As ListView, PathArchivo _
As String, sChar As String)
 On Error GoTo errsub
 Dim Linea As String, x As Integer, i As Integer
 Open PathArchivo For Output As #1
 With ListView
 For i = 1 To ListView.ListItems.Count
 Linea = .ListItems(i) & sChar
 For x = 1 To ListView.ColumnHeaders.Count - 1
 Linea = Linea & .ListItems.Item(i).SubItems(x) & sChar
 Next
 Print #1, Linea
 Next
 Close
 MsgBox " Archivo de Texto generado en: " & vbCrLf _
 & PathArchivo, vbInformation
 End With
 Exit Sub
errsub:
 MsgBox Err.Description, vbCritical
 Close
End Sub

Public Sub Exportar_ListViewExel()
 On Error GoTo nose
 Call FileCopy(App.Path & "\archivo.sys", Text2.Text)
nose:
End Sub

Private Sub Text1_Change(Index As Integer)
 Select Case Index
 Case (0)
 If Text1(0).Text = "" Then
 cmdExportar.Enabled = False
 Else
 cmdExportar.Enabled = True
 End If
 Case (1)
 If Text1(1).Text = "" Then
 cmdExportarHTML.Enabled = False
 Else
 cmdExportarHTML.Enabled = True
 End If
 End Select
End Sub

Private Sub Text2_Change()
 If Text2.Text = "" Then
 cmdexportarexel.Enabled = False
 Else
 cmdexportarexel.Enabled = True
 End If
End Sub
