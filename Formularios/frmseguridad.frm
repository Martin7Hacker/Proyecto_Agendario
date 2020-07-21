VERSION 5.00
Begin VB.Form frmseguridad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seguridad por Contraceña"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8370
   Icon            =   "frmseguridad.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5385
   ScaleWidth      =   8370
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton Option1 
      Caption         =   "Cada contraseña dependiente del usuario puede ser diferente o iguales"
      Height          =   615
      Index           =   2
      Left            =   5160
      TabIndex        =   70
      Top             =   4080
      Width           =   2895
   End
   Begin VB.PictureBox Picture18 
      Height          =   3495
      Left            =   8270
      ScaleHeight     =   3435
      ScaleWidth      =   0
      TabIndex        =   69
      Top             =   360
      Width           =   60
   End
   Begin VB.PictureBox Picture17 
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3435
      ScaleWidth      =   0
      TabIndex        =   68
      Top             =   360
      Width           =   60
   End
   Begin VB.PictureBox Picture13 
      Height          =   60
      Left            =   120
      ScaleHeight     =   0
      ScaleWidth      =   8115
      TabIndex        =   65
      Top             =   3840
      Width           =   8175
   End
   Begin VB.PictureBox Picture1 
      Height          =   60
      Left            =   120
      ScaleHeight     =   0
      ScaleWidth      =   8115
      TabIndex        =   64
      Top             =   360
      Width           =   8175
      Begin VB.PictureBox Picture16 
         Height          =   1815
         Left            =   -1440
         ScaleHeight     =   1755
         ScaleWidth      =   1395
         TabIndex        =   67
         Top             =   -1800
         Width           =   1455
      End
      Begin VB.PictureBox Picture15 
         Height          =   1815
         Left            =   -1440
         ScaleHeight     =   1755
         ScaleWidth      =   1395
         TabIndex        =   66
         Top             =   -1800
         Width           =   1455
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3225
      Left            =   7990
      Max             =   -2
      TabIndex        =   63
      Top             =   495
      Width           =   250
   End
   Begin VB.PictureBox Picture14 
      BorderStyle     =   0  'None
      Height          =   3495
      Left            =   240
      ScaleHeight     =   3495
      ScaleWidth      =   8055
      TabIndex        =   6
      Top             =   360
      Width           =   8055
      Begin VB.PictureBox picDesplasar 
         BorderStyle     =   0  'None
         Height          =   6375
         Left            =   0
         ScaleHeight     =   6375
         ScaleWidth      =   8775
         TabIndex        =   7
         Top             =   0
         Width           =   8775
         Begin VB.CommandButton cmdpoderImprimir 
            Height          =   255
            Left            =   7320
            Picture         =   "frmseguridad.frx":0CCA
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   4410
            Width           =   255
         End
         Begin VB.TextBox Text11 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   3240
            Locked          =   -1  'True
            PasswordChar    =   "*"
            TabIndex        =   50
            Top             =   4870
            Width           =   3975
         End
         Begin VB.PictureBox Picture12 
            Height          =   60
            Left            =   0
            ScaleHeight     =   0
            ScaleWidth      =   7515
            TabIndex        =   49
            Top             =   4770
            Width           =   7575
         End
         Begin VB.CommandButton cmdcopiaseguridad 
            Height          =   255
            Left            =   7320
            Picture         =   "frmseguridad.frx":1994
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   4860
            Width           =   255
         End
         Begin VB.TextBox Text10 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   3240
            Locked          =   -1  'True
            PasswordChar    =   "*"
            TabIndex        =   47
            Top             =   4410
            Width           =   3975
         End
         Begin VB.PictureBox Picture11 
            Height          =   60
            Left            =   0
            ScaleHeight     =   0
            ScaleWidth      =   7515
            TabIndex        =   46
            Top             =   5220
            Width           =   7575
         End
         Begin VB.PictureBox Picture10 
            Height          =   60
            Left            =   0
            ScaleHeight     =   0
            ScaleWidth      =   7515
            TabIndex        =   45
            Top             =   4290
            Width           =   7575
         End
         Begin VB.TextBox Text9 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   3240
            Locked          =   -1  'True
            PasswordChar    =   "*"
            TabIndex        =   44
            Top             =   3930
            Width           =   3975
         End
         Begin VB.CommandButton cmdexportar 
            Height          =   255
            Left            =   7320
            Picture         =   "frmseguridad.frx":265E
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   3930
            Width           =   255
         End
         Begin VB.PictureBox Picture9 
            Height          =   60
            Left            =   0
            ScaleHeight     =   0
            ScaleWidth      =   7515
            TabIndex        =   42
            Top             =   3840
            Width           =   7575
         End
         Begin VB.TextBox Text8 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   3240
            Locked          =   -1  'True
            PasswordChar    =   "*"
            TabIndex        =   41
            Top             =   3480
            Width           =   3975
         End
         Begin VB.CommandButton cmdelimniarSelecionado 
            Height          =   255
            Left            =   7320
            Picture         =   "frmseguridad.frx":3328
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   3480
            Width           =   255
         End
         Begin VB.PictureBox Picture8 
            Height          =   60
            Left            =   0
            ScaleHeight     =   0
            ScaleWidth      =   7515
            TabIndex        =   39
            Top             =   3360
            Width           =   7575
         End
         Begin VB.TextBox Text7 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   3240
            Locked          =   -1  'True
            PasswordChar    =   "*"
            TabIndex        =   38
            Top             =   3000
            Width           =   3975
         End
         Begin VB.CommandButton cmdElimniarTodo 
            Height          =   255
            Left            =   7320
            Picture         =   "frmseguridad.frx":3FF2
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   3000
            Width           =   255
         End
         Begin VB.CommandButton cmdiniciodelprograma 
            Height          =   255
            Left            =   7320
            Picture         =   "frmseguridad.frx":4CBC
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   2520
            Width           =   255
         End
         Begin VB.CommandButton cmdPoderverlaPlanilla 
            Height          =   255
            Left            =   7320
            Picture         =   "frmseguridad.frx":5986
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   2040
            Width           =   255
         End
         Begin VB.CommandButton cmdbusquedaarchivos 
            Height          =   255
            Left            =   7320
            Picture         =   "frmseguridad.frx":6650
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   1560
            Width           =   255
         End
         Begin VB.CommandButton cmdcreacionderegistros 
            Height          =   255
            Left            =   7320
            Picture         =   "frmseguridad.frx":731A
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   1080
            Width           =   255
         End
         Begin VB.CommandButton cmdmodificaciondedatos 
            Height          =   255
            Left            =   7320
            Picture         =   "frmseguridad.frx":7FE4
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton cmdcontracenia 
            Height          =   255
            Left            =   7320
            Picture         =   "frmseguridad.frx":8CAE
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   120
            Width           =   255
         End
         Begin VB.TextBox Text6 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   3240
            Locked          =   -1  'True
            PasswordChar    =   "*"
            TabIndex        =   30
            Top             =   2520
            Width           =   3975
         End
         Begin VB.PictureBox Picture7 
            Height          =   60
            Left            =   0
            ScaleHeight     =   0
            ScaleWidth      =   7515
            TabIndex        =   29
            Top             =   2880
            Width           =   7575
         End
         Begin VB.TextBox Text5 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   3240
            Locked          =   -1  'True
            PasswordChar    =   "*"
            TabIndex        =   28
            Top             =   2040
            Width           =   3975
         End
         Begin VB.PictureBox Picture6 
            Height          =   60
            Left            =   0
            ScaleHeight     =   0
            ScaleWidth      =   7515
            TabIndex        =   27
            Top             =   2400
            Width           =   7575
         End
         Begin VB.TextBox Text4 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   3240
            Locked          =   -1  'True
            PasswordChar    =   "*"
            TabIndex        =   26
            Top             =   1560
            Width           =   3975
         End
         Begin VB.PictureBox Picture5 
            Height          =   60
            Left            =   0
            ScaleHeight     =   0
            ScaleWidth      =   7515
            TabIndex        =   25
            Top             =   1920
            Width           =   7575
         End
         Begin VB.TextBox Text3 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   3240
            Locked          =   -1  'True
            PasswordChar    =   "*"
            TabIndex        =   24
            Top             =   1080
            Width           =   3975
         End
         Begin VB.PictureBox Picture4 
            Height          =   60
            Left            =   0
            ScaleHeight     =   0
            ScaleWidth      =   7515
            TabIndex        =   23
            Top             =   1440
            Width           =   7575
         End
         Begin VB.TextBox Text2 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   3240
            Locked          =   -1  'True
            PasswordChar    =   "*"
            TabIndex        =   22
            Top             =   600
            Width           =   3975
         End
         Begin VB.PictureBox Picture3 
            Height          =   60
            Left            =   0
            ScaleHeight     =   0
            ScaleWidth      =   7515
            TabIndex        =   21
            Top             =   960
            Width           =   7575
         End
         Begin VB.TextBox Text1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   3240
            Locked          =   -1  'True
            PasswordChar    =   "*"
            TabIndex        =   20
            Top             =   120
            Width           =   3975
         End
         Begin VB.PictureBox Picture2 
            Height          =   60
            Left            =   0
            ScaleHeight     =   0
            ScaleWidth      =   7515
            TabIndex        =   19
            Top             =   480
            Width           =   7575
         End
         Begin VB.CommandButton cmdborrar 
            Height          =   255
            Index           =   0
            Left            =   840
            Picture         =   "frmseguridad.frx":9978
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Borrar Contraceña"
            Top             =   120
            Width           =   255
         End
         Begin VB.CommandButton cmdborrar 
            Height          =   255
            Index           =   1
            Left            =   840
            Picture         =   "frmseguridad.frx":A642
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Borrar Contraceña"
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton cmdborrar 
            Height          =   255
            Index           =   2
            Left            =   840
            Picture         =   "frmseguridad.frx":B30C
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Borrar Contraceña"
            Top             =   1080
            Width           =   255
         End
         Begin VB.CommandButton cmdborrar 
            Height          =   255
            Index           =   3
            Left            =   840
            Picture         =   "frmseguridad.frx":BFD6
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Borrar Contraceña"
            Top             =   1560
            Width           =   255
         End
         Begin VB.CommandButton cmdborrar 
            Height          =   255
            Index           =   4
            Left            =   840
            Picture         =   "frmseguridad.frx":CCA0
            Style           =   1  'Graphical
            TabIndex        =   14
            ToolTipText     =   "Borrar Contraceña"
            Top             =   2040
            Width           =   255
         End
         Begin VB.CommandButton cmdborrar 
            Height          =   255
            Index           =   5
            Left            =   840
            Picture         =   "frmseguridad.frx":D96A
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Borrar Contraceña"
            Top             =   2520
            Width           =   255
         End
         Begin VB.CommandButton cmdborrar 
            Height          =   255
            Index           =   6
            Left            =   840
            Picture         =   "frmseguridad.frx":E634
            Style           =   1  'Graphical
            TabIndex        =   12
            ToolTipText     =   "Borrar Contraceña"
            Top             =   3000
            Width           =   255
         End
         Begin VB.CommandButton cmdborrar 
            Height          =   255
            Index           =   7
            Left            =   840
            Picture         =   "frmseguridad.frx":F2FE
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Borrar Contraceña"
            Top             =   3480
            Width           =   255
         End
         Begin VB.CommandButton cmdborrar 
            Height          =   255
            Index           =   8
            Left            =   840
            Picture         =   "frmseguridad.frx":FFC8
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Borrar Contraceña"
            Top             =   3960
            Width           =   255
         End
         Begin VB.CommandButton cmdborrar 
            Height          =   255
            Index           =   9
            Left            =   840
            Picture         =   "frmseguridad.frx":10C92
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Borrar Contraceña"
            Top             =   4440
            Width           =   255
         End
         Begin VB.CommandButton cmdborrar 
            Height          =   255
            Index           =   10
            Left            =   840
            Picture         =   "frmseguridad.frx":1195C
            Style           =   1  'Graphical
            TabIndex        =   8
            ToolTipText     =   "Borrar Contraceña"
            Top             =   4920
            Width           =   255
         End
         Begin VB.Image Image11 
            Height          =   480
            Left            =   0
            Picture         =   "frmseguridad.frx":12626
            Top             =   4320
            Width           =   480
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Poder Imprimir:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2160
            TabIndex        =   62
            Top             =   4460
            Width           =   1035
         End
         Begin VB.Image Image10 
            Height          =   480
            Left            =   0
            Picture         =   "frmseguridad.frx":132F0
            Top             =   4890
            Width           =   480
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Crear copia de Seguridad:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1320
            TabIndex        =   61
            Top             =   4930
            Width           =   1845
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "poder exportar Archivos:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1440
            TabIndex        =   60
            Top             =   3990
            Width           =   1725
         End
         Begin VB.Image Image9 
            Height          =   480
            Left            =   0
            Picture         =   "frmseguridad.frx":13FBA
            Top             =   3840
            Width           =   480
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Eliminar Seleciónado:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1680
            TabIndex        =   59
            Top             =   3530
            Width           =   1515
         End
         Begin VB.Image Image8 
            Height          =   480
            Left            =   0
            Picture         =   "frmseguridad.frx":14C84
            Top             =   3420
            Width           =   480
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Eliminar todo:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   2160
            TabIndex        =   58
            Top             =   3015
            Width           =   945
         End
         Begin VB.Image Image7 
            Height          =   480
            Left            =   0
            Picture         =   "frmseguridad.frx":1594E
            Top             =   2880
            Width           =   480
         End
         Begin VB.Image Image6 
            Height          =   480
            Left            =   0
            Picture         =   "frmseguridad.frx":16618
            Top             =   0
            Width           =   480
         End
         Begin VB.Image Image5 
            Height          =   480
            Left            =   0
            Picture         =   "frmseguridad.frx":172E2
            Top             =   2520
            Width           =   480
         End
         Begin VB.Image Image4 
            Height          =   480
            Left            =   0
            Picture         =   "frmseguridad.frx":17FAC
            Top             =   1920
            Width           =   480
         End
         Begin VB.Image Image3 
            Height          =   480
            Left            =   0
            Picture         =   "frmseguridad.frx":18C76
            Top             =   1440
            Width           =   480
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   0
            Picture         =   "frmseguridad.frx":19940
            Top             =   1080
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   0
            Picture         =   "frmseguridad.frx":1A60A
            Top             =   600
            Width           =   480
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "pers. / Ver contraceñas:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1320
            TabIndex        =   57
            Top             =   2550
            Width           =   1725
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Poder ver la Planilla Virtual:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1200
            TabIndex        =   56
            Top             =   2090
            Width           =   1920
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "&Busqueda de Archivos:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1440
            TabIndex        =   55
            Top             =   1610
            Width           =   1650
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Creación de Registros:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1440
            TabIndex        =   54
            Top             =   1130
            Width           =   1605
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Modificación de Datos:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1440
            TabIndex        =   53
            Top             =   630
            Width           =   1635
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio de Seguridad:"
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   1680
            TabIndex        =   52
            Top             =   130
            Width           =   1410
         End
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Sin protección de Contraseña"
      Height          =   615
      Index           =   1
      Left            =   2880
      TabIndex        =   5
      Top             =   4080
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Permitir todas las contraceñas Unitarias."
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   4
      Top             =   4080
      Width           =   2295
   End
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   100
      TabIndex        =   1
      Top             =   3840
      Width           =   8175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Seguridad de control por contraseña de Agendario v1.0"
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
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   4740
   End
End
Attribute VB_Name = "frmseguridad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Open Source
'* System Application Software
'* Programa frmseguridad de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************

Private Sub cmdaceptar_Click()
 guardarArchivos
 Unload Me
End Sub

Private Sub cmdborrar_Click(Index As Integer)
 Select Case MsgBox("Quieres eliminar la contraceña", _
 vbExclamation + vbYesNo, nombre_programa)
 Case (vbYes)
 Select Case Index
 Case (0): Text1.Text = ""
 Case (1): Text2.Text = ""
 Case (2): Text3.Text = ""
 Case (3): Text4.Text = ""
 Case (4): Text5.Text = ""
 Case (5): Text6.Text = ""
 Case (6): Text7.Text = ""
 Case (7): Text8.Text = ""
 Case (8): Text9.Text = ""
 Case (9): Text10.Text = ""
 Case (10): Text11.Text = ""
 End Select
 End Select
End Sub

Private Sub cmdbusquedaarchivos_Click()
 controlPasarContracena Text4.Text, 3
 frmcontracena.Show 1
End Sub

Private Sub cmdcancelar_Click()
 Unload Me
End Sub

Private Sub cmdcontracenia_Click()
 controlPasarContracena Text1.Text, 0
 frmcontracena.Show 1
End Sub

Public Sub controlPasarContracena(ByRef contra As String, ByVal control _
As Integer)
 frmcontracena.control_fx = control
Select Case control
Case 0
If contra = "" Then
       
       contra = Text1.Text
      
       Else
         Text1.Text = contra
          With frmcontracena
            .Text1 = Text1.Text
            .Text2 = Text1.Text
       End With
       End If
       
       Case 1
       
        If contra = "" Then
       contra = Text2.Text
       
       Else
       
         Text2.Text = contra
         With frmcontracena
            .Text1 = Text2.Text
            .Text2 = Text2.Text
       End With
       End If
       
       Case 2
       
        If contra = "" Then
       contra = Text3.Text
       
       Else
         Text3.Text = contra
         With frmcontracena
            .Text1 = Text3.Text
            .Text2 = Text3.Text
       End With
       End If
       
       Case 3
       
        If contra = "" Then
       contra = Text4.Text
       
       Else
         Text4.Text = contra
         With frmcontracena
            .Text1 = Text4.Text
            .Text2 = Text4.Text
       End With
       End If
       
       Case 4
       
        If contra = "" Then
       contra = Text5.Text
       
       Else
         Text5.Text = contra
         With frmcontracena
            .Text1 = Text5.Text
            .Text2 = Text5.Text
       End With
       End If
       
       Case 5
       
        If contra = "" Then
       contra = Text6.Text
       
       Else
         Text6.Text = contra
         With frmcontracena
            .Text1 = Text6.Text
            .Text2 = Text6.Text
       End With
       End If
       
        Case 6
       
        If contra = "" Then
       contra = Text7.Text
       
       Else
       
         Text7.Text = contra
         With frmcontracena
            .Text1 = Text7.Text
            .Text2 = Text7.Text
       End With
       End If
       
        Case 7
       
        If contra = "" Then
       contra = Text8.Text
       
       Else
       
         Text8.Text = contra
         With frmcontracena
            .Text1 = Text8.Text
            .Text2 = Text8.Text
       End With
       End If
       
        Case 8
       
        If contra = "" Then
       contra = Text9.Text
       
       Else
       
         Text9.Text = contra
         With frmcontracena
            .Text1 = Text9.Text
            .Text2 = Text9.Text
       End With
       End If
       
          Case 9
       
        If contra = "" Then
       contra = Text10.Text
       
       Else
       
         Text10.Text = contra
         With frmcontracena
            .Text1 = Text10.Text
            .Text2 = Text10.Text
       End With
       End If
       
          Case 10
       
        If contra = "" Then
       contra = Text11.Text
       
       Else
       
         Text11.Text = contra
         
         With frmcontracena
            .Text1 = Text11.Text
            .Text2 = Text11.Text
       End With
       End If
       
End Select
End Sub

Private Sub cmdcopiaseguridad_Click()
 controlPasarContracena Text11.Text, 10
 frmcontracena.Show 1
End Sub

Private Sub cmdcreacionderegistros_Click()
 controlPasarContracena Text3.Text, 2
 frmcontracena.Show 1
End Sub

Private Sub cmdelimniarSelecionado_Click()
 controlPasarContracena Text8.Text, 7
 frmcontracena.Show 1
End Sub

Private Sub cmdElimniarTodo_Click()
 controlPasarContracena Text7.Text, 6
 frmcontracena.Show 1
End Sub

Private Sub cmdexportar_Click()
 controlPasarContracena Text9.Text, 8
 frmcontracena.Show 1
End Sub

Private Sub cmdiniciodelprograma_Click()
 controlPasarContracena Text6.Text, 5
 frmcontracena.Show 1
End Sub

Private Sub cmdmodificaciondedatos_Click()
 controlPasarContracena Text2.Text, 1
 frmcontracena.Show 1
End Sub

Private Sub cmdpoderImprimir_Click()
 controlPasarContracena Text10.Text, 9
 frmcontracena.Show 1
End Sub

Private Sub cmdPoderverlaPlanilla_Click()
 controlPasarContracena Text5.Text, 4
 frmcontracena.Show 1
End Sub

Private Sub unificarcontracenia()
 cmdcontracenia_Click
End Sub

Private Sub Form_Load()
 Text1.Text = y_seguridad.y_inicio
 Text2.Text = y_seguridad.y_modific
 Text3.Text = y_seguridad.y_creacion
 Text4.Text = y_seguridad.y_busqueda
 Text5.Text = y_seguridad.y_poderver
 Text6.Text = y_seguridad.y_iniciodel
 Text7.Text = y_seguridad.y_elimnarTodo
 Text8.Text = y_seguridad.y_eliminarSeleccionado
 Text9.Text = y_seguridad.y_poderExportar
 Text10.Text = y_seguridad.y_eliminarSeleccionado
 Text11.Text = y_seguridad.y_poderExportar
End Sub

Private Sub Option1_Click(Index As Integer)

Select Case Index

       Case 0
       
        EnablesEstado False
        unificarcontracenia
        EnablesTexto True
       
       Case 1
       
       Select Case MsgBox("Quieres utilizar este programa sin contraceña?", _
       vbExclamation + vbYesNo, nombre_programa)
       
              Case vbYes
              Text1.Text = ""
              Text2.Text = ""
              Text3.Text = ""
              Text4.Text = ""
              Text5.Text = ""
              Text6.Text = ""
              Text7.Text = ""
              Text8.Text = ""
              Text9.Text = ""
              Text10.Text = ""
              Text11.Text = ""
              EnablesEstado False
              EnablesTexto False
              
        End Select
        
        Case 2
     
          EnablesEstado True
          EnablesTexto True

End Select
   
End Sub

Private Sub EnablesEstado(ByVal estado As Boolean)
 cmdcontracenia.Enabled = estado
 cmdmodificaciondedatos.Enabled = estado
 cmdcreacionderegistros.Enabled = estado
 cmdbusquedaarchivos.Enabled = estado
 cmdPoderverlaPlanilla.Enabled = estado
 cmdiniciodelprograma.Enabled = estado
 cmdcopiaseguridad.Enabled = estado
 cmdpoderImprimir.Enabled = estado
 cmdelimniarSelecionado.Enabled = estado
 cmdElimniarTodo.Enabled = estado
 cmdexportar.Enabled = estado
End Sub

Private Sub EnablesTexto(ByVal estado As Boolean)
 Text1.Enabled = estado
 Text2.Enabled = estado
 Text3.Enabled = estado
 Text4.Enabled = estado
 Text5.Enabled = estado
 Text6.Enabled = estado
 Text7.Enabled = estado
 Text8.Enabled = estado
 Text9.Enabled = estado
 Text10.Enabled = estado
 Text11.Enabled = estado
End Sub

Private Sub guardarArchivos()
 y_seguridad.y_inicio = Text1.Text
 y_seguridad.y_modific = Text2.Text
 y_seguridad.y_creacion = Text3.Text
 y_seguridad.y_busqueda = Text4.Text
 y_seguridad.y_poderver = Text5.Text
 y_seguridad.y_iniciodel = Text6.Text
 y_seguridad.y_elimnarTodo = Text7.Text
 y_seguridad.y_eliminarSeleccionado = Text8.Text
 y_seguridad.y_poderExportar = Text9.Text
 y_seguridad.y_poderImprimir = Text10.Text
 y_seguridad.y_crearCopiaSeguridad = Text11.Text
 ModPrincipal.Guardar
End Sub

Private Sub VScroll1_Change()
 desplazar
End Sub

Private Sub VScroll1_Scroll()
 desplazar
End Sub

Private Sub desplazar()
 picDesplasar.Top = VScroll1.Value * 900
End Sub
