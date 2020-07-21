VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmbusqueda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búscar en Agendario v1.0"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7695
   Icon            =   "frmbusqueda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   7695
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   2760
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H8000000F&
      ForeColor       =   &H00800080&
      Height          =   315
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1950
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "H:mm:ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3082
         SubFormatType   =   0
      EndProperty
      Height          =   345
      Left            =   6090
      TabIndex        =   12
      Top             =   435
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      _Version        =   393216
      MouseIcon       =   "frmbusqueda.frx":0CCA
      CalendarBackColor=   -2147483633
      CalendarTitleBackColor=   8388736
      Format          =   49610753
      CurrentDate     =   42318
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   375
      Left            =   5640
      TabIndex        =   11
      Top             =   1320
      Width           =   255
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
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
      Height          =   405
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "10"
      Top             =   1320
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Permitir Búsqueda Inteligente."
      ForeColor       =   &H00800080&
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   960
      Value           =   1  'Checked
      Width           =   3255
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   2880
      Top             =   2160
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdbuscar 
      Caption         =   "Búscar"
      Height          =   855
      Left            =   6480
      Picture         =   "frmbusqueda.frx":19A4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox bdato 
      BackColor       =   &H8000000F&
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
      Height          =   495
      Left            =   2760
      TabIndex        =   4
      Top             =   360
      Width           =   4815
   End
   Begin VB.PictureBox Picture2 
      Height          =   60
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   9675
      TabIndex        =   3
      Top             =   2640
      Width           =   9735
   End
   Begin VB.ListBox List1 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Height          =   60
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   9675
      TabIndex        =   0
      Top             =   240
      Width           =   9735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Búsqueda:"
      ForeColor       =   &H00800080&
      Height          =   435
      Left            =   3120
      TabIndex        =   15
      Top             =   1800
      Width           =   1020
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   5950
      Picture         =   "frmbusqueda.frx":266E
      Top             =   1250
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5950
      Picture         =   "frmbusqueda.frx":3338
      Top             =   1850
      Width           =   480
   End
   Begin VB.Label Label2 
      Caption         =   "Velocidad de Búsqueda:"
      ForeColor       =   &H00800080&
      Height          =   195
      Left            =   2760
      TabIndex        =   9
      Top             =   1395
      Width           =   1740
   End
   Begin VB.Label labbuscar 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Registros donde Búscar:"
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
      Top             =   0
      Width           =   7500
   End
End
Attribute VB_Name = "frmbusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Open Source
'* System Application Software
'* Programa Búsqueda de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************
Private Sub bdato_KeyPress(KeyAscii As Integer)
frmdatos.soloLetrasMayusculas 10, KeyAscii
End Sub

Private Sub cmdaceptar_Click()
Open "datos.txt" For Output As 1
Dim d As Integer
For d = 0 To 31
Print #1, List1.List(d)
Next
Close #1
Unload Me

End Sub

Private Sub cmdbuscar_Click()
Timer1.Interval = Text1.Text
Select Case Check1.Value
Case (0)
Timer1.Enabled = False
Case (1)
Timer1.Enabled = True
End Select
ModFunciones_Publicas.busqueda_virtual List1.ListIndex
End Sub

Private Sub cmdcancelar_Click()
Unload Me
End Sub

Private Sub Combo1_Change()
Select Case Combo1.ListIndex
       Case 0
       DTPicker1.Format = dtpShortDate
       control 0
       Case 1
       DTPicker1.Format = dtpTime
       control 1
End Select
End Sub

Private Sub Combo1_Click()
Combo1_Change
End Sub

Private Sub Combo1_Scroll()
Combo1_Click
End Sub

Private Sub control(ByVal ind As Integer)
 bdato.Text = ""
 Select Case (ind)
  Case (0)
     DTPicker1.Value = CDate(Date)
     bdato.Text = DTPicker1.Value
  Case (1)
      DTPicker1.Value = CDate(Time)
      bdato.Text = DTPicker1.Value
End Select
 bdato.Text = DTPicker1.Value
End Sub

Private Sub DTPicker1_Click()
 bdato.Text = DTPicker1.Value
End Sub

Private Sub Form_Load()
MDIPrincipal.Enabled = False
        Me.Icon = MDIPrincipal.Icon
        cargar_datos
        List1.ListIndex = 0
        VScroll1 = Text1.Text
        tipobusqueda
        Combo1.ListIndex = 0
End Sub

Private Sub tipobusqueda()
With Combo1
     .Clear
     .AddItem "Fecha"
     .AddItem "Hora"
End With
End Sub

Private Sub cargar_datos()
With List1
        List1.Clear
        Open "etiquetas.opxl" For Input As 1
        Dim etiqueta As String
        Do While Not EOF(1)
        Line Input #1, etiqueta
        .AddItem escriptar.funcion_desescriptar(etiqueta)
        Loop
        Close #1
        End With
        Me.Caption = "Búscar en " & nombre_programa
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIPrincipal.Enabled = True
Mod_Funciones_conByts.desoprimr_boton 4
End Sub

Private Sub List1_Click()
labbuscar.Caption = List1.List(List1.ListIndex) & " :"
End Sub

Private Sub Text1_Change()
VScroll1.Value = Text1.Text
End Sub

Private Sub Timer1_Timer()
On Error GoTo nose
    ModFunciones_Publicas.busqueda_virtual List1.ListIndex
nose:
End Sub

Private Sub VScroll1_Change()
Text1.Text = VScroll1.Value
End Sub
