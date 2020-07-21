VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmModificar 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Modificar Contacto"
   ClientHeight    =   6345
   ClientLeft      =   45
   ClientTop       =   675
   ClientWidth     =   9630
   Icon            =   "frmModificar.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6345
   ScaleWidth      =   9630
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   240
      ScaleHeight     =   4695
      ScaleWidth      =   6735
      TabIndex        =   14
      Top             =   1440
      Width           =   6735
      Begin VB.PictureBox picHoja 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   8535
         Left            =   240
         ScaleHeight     =   8535
         ScaleWidth      =   6735
         TabIndex        =   15
         Top             =   0
         Width           =   6735
         Begin VB.ComboBox cobEstado 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4800
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   7440
            Width           =   1560
         End
         Begin VB.CommandButton cmdlinkdefacebook 
            Height          =   255
            Left            =   6120
            Picture         =   "frmModificar.frx":0CCA
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   6240
            Width           =   255
         End
         Begin VB.CommandButton cmdlinkfacebook 
            Height          =   255
            Left            =   6120
            Picture         =   "frmModificar.frx":1994
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   6000
            Width           =   255
         End
         Begin VB.CommandButton cmdlinktuitter 
            Height          =   255
            Left            =   6120
            Picture         =   "frmModificar.frx":265E
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   6960
            Width           =   255
         End
         Begin VB.CommandButton cmdlinktuiter 
            Height          =   255
            Left            =   6120
            Picture         =   "frmModificar.frx":3328
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   6720
            Width           =   255
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   17
            Left            =   2520
            TabIndex        =   38
            Top             =   6000
            Width           =   3615
         End
         Begin VB.CommandButton cmdpaises 
            Height          =   255
            Left            =   6120
            Picture         =   "frmModificar.frx":3FF2
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   3150
            Width           =   255
         End
         Begin VB.ComboBox cobSexo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   4560
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   3840
            Width           =   1815
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   31
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   52
            Top             =   3840
            Width           =   2055
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   30
            Left            =   2520
            TabIndex        =   51
            Top             =   4080
            Width           =   3855
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   29
            Left            =   2520
            TabIndex        =   50
            Top             =   4320
            Width           =   3855
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   28
            Left            =   2520
            TabIndex        =   49
            Text            =   "@.COM"
            Top             =   4560
            Width           =   3855
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   27
            Left            =   2520
            TabIndex        =   48
            Text            =   "@.COM"
            Top             =   4800
            Width           =   3855
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   26
            Left            =   2520
            TabIndex        =   47
            Text            =   "@.COM"
            Top             =   5040
            Width           =   3855
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Index           =   25
            Left            =   2520
            MaxLength       =   3
            TabIndex        =   46
            Top             =   5280
            Width           =   3855
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   3082
               SubFormatType   =   3
            EndProperty
            Height          =   285
            Index           =   24
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   5520
            Width           =   2295
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   23
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   44
            Top             =   7440
            Width           =   3855
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   22
            Left            =   2520
            TabIndex        =   43
            Top             =   7200
            Width           =   3855
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   21
            Left            =   2520
            TabIndex        =   42
            Top             =   6960
            Width           =   3615
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   20
            Left            =   2520
            TabIndex        =   41
            Top             =   6720
            Width           =   3615
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   19
            Left            =   2520
            TabIndex        =   40
            Top             =   6480
            Width           =   3855
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   18
            Left            =   2520
            TabIndex        =   39
            Top             =   6240
            Width           =   3615
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   16
            Left            =   2520
            TabIndex        =   37
            Top             =   5760
            Width           =   3855
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   15
            Left            =   2520
            TabIndex        =   36
            Top             =   1680
            Width           =   3855
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   14
            Left            =   2520
            TabIndex        =   35
            Top             =   1920
            Width           =   3855
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   765
            Index           =   13
            Left            =   2520
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   34
            Top             =   2160
            Width           =   3855
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   12
            Left            =   2520
            TabIndex        =   33
            Top             =   2880
            Width           =   3855
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   11
            Left            =   6360
            TabIndex        =   32
            Top             =   3360
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   10
            Left            =   2520
            TabIndex        =   31
            Top             =   3120
            Width           =   3615
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   9
            Left            =   2520
            TabIndex        =   30
            Top             =   3360
            Width           =   3855
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   8
            Left            =   2520
            TabIndex        =   29
            Top             =   3600
            Width           =   3855
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   7
            Left            =   2520
            TabIndex        =   28
            Top             =   1440
            Width           =   3855
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   6
            Left            =   5520
            TabIndex        =   27
            Top             =   1440
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   5
            Left            =   2520
            TabIndex        =   26
            Top             =   1200
            Width           =   3855
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   4
            Left            =   2520
            TabIndex        =   25
            Top             =   960
            Width           =   3855
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   3
            Left            =   2520
            TabIndex        =   24
            Top             =   720
            Width           =   3855
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   2520
            TabIndex        =   23
            Top             =   480
            Width           =   3855
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   2520
            TabIndex        =   22
            Top             =   240
            Width           =   3855
         End
         Begin VB.TextBox txtnombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   0
            Left            =   2520
            TabIndex        =   21
            Top             =   0
            Width           =   3855
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   285
            Left            =   4800
            TabIndex        =   55
            Top             =   5520
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   503
            _Version        =   393216
            CalendarTitleBackColor=   8388736
            Format          =   49938433
            CurrentDate     =   42318
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sexo:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   31
            Left            =   120
            TabIndex        =   100
            Top             =   3840
            Width           =   2325
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Calle:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   30
            Left            =   120
            TabIndex        =   99
            Top             =   4080
            Width           =   2295
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Calle:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   29
            Left            =   120
            TabIndex        =   98
            Top             =   4320
            Width           =   2295
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   28
            Left            =   120
            TabIndex        =   97
            Top             =   4560
            Width           =   2340
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   27
            Left            =   120
            TabIndex        =   96
            Top             =   4800
            Width           =   2340
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Email:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   26
            Left            =   120
            TabIndex        =   95
            Top             =   5040
            Width           =   2340
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Edad:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   25
            Left            =   120
            TabIndex        =   94
            Top             =   5280
            Width           =   2220
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha N:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   24
            Left            =   120
            TabIndex        =   93
            Top             =   5520
            Width           =   2220
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "E.Civil:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   23
            Left            =   120
            TabIndex        =   92
            Top             =   7440
            Width           =   2400
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N Casas:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   22
            Left            =   120
            TabIndex        =   91
            Top             =   7200
            Width           =   2325
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2 Link de Twitter:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   21
            Left            =   120
            TabIndex        =   90
            Top             =   6960
            Width           =   2295
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Link de Twitter:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   89
            Top             =   6720
            Width           =   2295
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "twitter:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   19
            Left            =   120
            TabIndex        =   88
            Top             =   6480
            Width           =   2385
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2 Link de Facebook:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   18
            Left            =   120
            TabIndex        =   87
            Top             =   6240
            Width           =   2295
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Link de Facebook:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   17
            Left            =   120
            TabIndex        =   86
            Top             =   5760
            Width           =   2295
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Facebook:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   16
            Left            =   120
            TabIndex        =   85
            Top             =   6000
            Width           =   2325
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Celular:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   84
            Top             =   1680
            Width           =   2325
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CI:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   83
            Top             =   1920
            Width           =   2475
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comentario:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   82
            Top             =   2160
            Width           =   2280
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   81
            Top             =   2880
            Width           =   2400
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   11
            Left            =   5400
            TabIndex        =   80
            Top             =   3360
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pais:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   79
            Top             =   3120
            Width           =   2385
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Departamento:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   78
            Top             =   3360
            Width           =   2250
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ciudad:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   77
            Top             =   3600
            Width           =   2340
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Celular:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   76
            Top             =   1440
            Width           =   2325
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telèfono:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   5520
            TabIndex        =   75
            Top             =   1440
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Telefono :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   74
            Top             =   1200
            Width           =   2400
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Teléfono :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   73
            Top             =   960
            Width           =   2400
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido P:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   72
            Top             =   720
            Width           =   2445
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apellido M:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   71
            ToolTipText     =   "Apellido Materno:"
            Top             =   480
            Width           =   2460
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "S Nombre:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   70
            ToolTipText     =   "Segundo Nombre:"
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   69
            ToolTipText     =   "Primer Nombre"
            Top             =   0
            Width           =   2400
         End
         Begin VB.Label lblverificar 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   1
            Left            =   0
            TabIndex        =   68
            Top             =   0
            Width           =   90
         End
         Begin VB.Label lblverificar 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   2
            Left            =   0
            TabIndex        =   67
            Top             =   480
            Width           =   90
         End
         Begin VB.Label lblverificar 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   3
            Left            =   0
            TabIndex        =   66
            Top             =   960
            Width           =   90
         End
         Begin VB.Label lblverificar 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   4
            Left            =   0
            TabIndex        =   65
            Top             =   1680
            Width           =   90
         End
         Begin VB.Label lblverificar 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   5
            Left            =   0
            TabIndex        =   64
            Top             =   1920
            Width           =   90
         End
         Begin VB.Label lblverificar 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   6
            Left            =   0
            TabIndex        =   63
            Top             =   2880
            Width           =   90
         End
         Begin VB.Label lblverificar 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   7
            Left            =   0
            TabIndex        =   62
            Top             =   3840
            Width           =   90
         End
         Begin VB.Label lblverificar 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   8
            Left            =   0
            TabIndex        =   61
            Top             =   4560
            Width           =   90
         End
         Begin VB.Label lblverificar 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   9
            Left            =   0
            TabIndex        =   60
            Top             =   4080
            Width           =   90
         End
         Begin VB.Label lblverificar 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   10
            Left            =   0
            TabIndex        =   59
            Top             =   5520
            Width           =   90
         End
         Begin VB.Label lblverificar 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   11
            Left            =   0
            TabIndex        =   58
            Top             =   7200
            Width           =   90
         End
         Begin VB.Label lblverificar 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   12
            Left            =   0
            TabIndex        =   57
            Top             =   7440
            Width           =   90
         End
         Begin VB.Label lblverificar 
            BackStyle       =   0  'Transparent
            Caption         =   "*"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   0
            Left            =   0
            TabIndex        =   56
            Top             =   720
            Width           =   90
         End
      End
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   4695
      Left            =   6960
      Max             =   -5
      TabIndex        =   13
      Top             =   1440
      Width           =   255
   End
   Begin VB.Timer Timer2 
      Interval        =   100
      Left            =   1800
      Top             =   4200
   End
   Begin VB.CommandButton cmdtomarhoradelsistema 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6960
      Picture         =   "frmModificar.frx":4CBC
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1119
      Width           =   255
   End
   Begin VB.CommandButton cmdcancelar 
      Appearance      =   0  'Flat
      Caption         =   "&Cancelar [F3]"
      Height          =   495
      Index           =   5
      Left            =   7440
      TabIndex        =   3
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   2640
   End
   Begin VB.CommandButton cmdModificar 
      Appearance      =   0  'Flat
      Caption         =   "&Permitir Edición [F1]"
      Height          =   495
      Index           =   3
      Left            =   7440
      TabIndex        =   1
      Top             =   3480
      Width           =   2055
   End
   Begin VB.CommandButton cmdModificar 
      Appearance      =   0  'Flat
      Caption         =   "&Modificar &Hoja de Agenda [F2]"
      Height          =   495
      Index           =   0
      Left            =   7440
      TabIndex        =   0
      Top             =   4080
      Width           =   2055
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2010
      Left            =   7230
      TabIndex        =   2
      Top             =   1200
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   3545
      _Version        =   393216
      ForeColor       =   8388736
      BackColor       =   255
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   49938434
      TitleBackColor  =   8388736
      TitleForeColor  =   -2147483639
      TrailingForeColor=   4210752
      CurrentDate     =   42327
   End
   Begin MSComCtl2.DTPicker DTPickerFechaHoy 
      Height          =   345
      Left            =   1800
      TabIndex        =   4
      Top             =   9120
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   609
      _Version        =   393216
      CalendarTitleBackColor=   8388736
      Format          =   49938433
      CurrentDate     =   42318
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   300
      Left            =   3000
      TabIndex        =   101
      Top             =   1125
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   16777215
      CalendarTitleBackColor=   16777215
      Format          =   49938434
      UpDown          =   -1  'True
      CurrentDate     =   0.805555555555556
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800080&
      BorderWidth     =   2
      Index           =   3
      X1              =   0
      X2              =   0
      Y1              =   1080
      Y2              =   6240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800080&
      BorderWidth     =   2
      Index           =   2
      X1              =   0
      X2              =   9600
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800080&
      BorderWidth     =   2
      Index           =   1
      X1              =   9600
      X2              =   9600
      Y1              =   6240
      Y2              =   1080
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Hora a Agendar:"
      Height          =   435
      Left            =   600
      TabIndex        =   102
      Top             =   1080
      Width           =   2370
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800080&
      BorderWidth     =   2
      Index           =   0
      X1              =   0
      X2              =   9600
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblfecha 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   50.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   1170
      Left            =   4320
      TabIndex        =   11
      Top             =   0
      Width           =   1110
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label lblmes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enero:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   270
      Left            =   6120
      TabIndex        =   9
      Top             =   240
      Width           =   705
   End
   Begin VB.Label lbldia 
      BackStyle       =   0  'Transparent
      Caption         =   "&Miercoles"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   315
      Left            =   6120
      TabIndex        =   8
      Top             =   600
      Width           =   1785
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "&Hora:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   360
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   1545
   End
   Begin VB.Label lblhora 
      BackStyle       =   0  'Transparent
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   840
      TabIndex        =   6
      Top             =   795
      Width           =   1440
   End
   Begin VB.Label lblanio 
      BackStyle       =   0  'Transparent
      Caption         =   "00000000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   1815
   End
   Begin VB.Menu Opciones 
      Caption         =   "Opciones"
      Begin VB.Menu permitiredicion 
         Caption         =   "&Permitir Edición"
         Shortcut        =   {F1}
      End
      Begin VB.Menu ModificarHojadeAgenda 
         Caption         =   "&Modificar Hoja de Agenda"
         Shortcut        =   {F2}
      End
      Begin VB.Menu cancelar 
         Caption         =   "&Cancelar"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "frmModificar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Open Source
'* System Application Software
'* Programa frmModificar de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************
Dim control(30), modx As Boolean: Dim textos(32) As String

Private Sub funcion_coloriar_controles(ByVal r As Integer, ByVal color _
As String)
 Dim i As Integer
 For i = 0 To txtnombre.Count - 1
 txtnombre(i).BackColor = &HFFFFFF
 txtnombre(i).Font.Bold = False 'le saco la negrita
 Label1(i).ForeColor = &HC000C0
 Next
 txtnombre(r).BackColor = color
 txtnombre(r).Font.Bold = True ' Establesco la letra en negrita
 Label1(r).ForeColor = color
End Sub

Private Sub cargarEstados()
 cobEstado.Clear
 Select Case cobSexo.ListIndex
 Case (0)
 With cobEstado
 .AddItem UCase("Single - Soltero")
 .AddItem UCase("Married - Casado")
 .AddItem UCase("Divorced - Divorciado")
 .AddItem UCase("Widower - Viudo")
 End With
 Case (1)
 With cobEstado
 .AddItem UCase("Single - Soltera")
 .AddItem UCase("Married - Casada ")
 .AddItem UCase("Divorcee -Divorciada")
 .AddItem UCase("Widow - Viuda")
 End With
 End Select
End Sub

Public Sub tipo_civil()
 txtnombre.Item(23).Text = UCase(cobEstado.Text)
 cargarEstados
End Sub

Private Sub cancelar_Click()
 Unload Me
End Sub

Private Sub cmdcancelar_Click(Index As Integer)
 Unload Me
End Sub

Private Sub cmdModificar_Click(Index As Integer)
 Select Case Index
 Case (0)
 Select Case MsgBox("Quieres Modificar este Registro", _
 vbYesNo + vbInformation, nombre_programa)
 Case (vbYes)
 funcion_modificar_datos
 End Select
 Case (1)
 Unload Me
 Case (2)
 funcion_visualizar_datos_enMemoria
 Case (3)
 If modx = False Then
 editar True
 modx = True
 cmdModificar(3).Caption = "&No Edición [F1]"
 ElseIf modx = True Then
 editar False
 modx = False
 cmdModificar(3).Caption = "&Permitir Edición [F1]"
 End If
 End Select
End Sub

Private Sub cmdpaises_Click()
 frmAplicarPaises.Show 1
End Sub

Private Sub cmdtomarhoradelsistema_Click()
 DTPicker2.Value = Time
 lblhora.Caption = DTPicker2.Value
End Sub

Private Sub cobEstado_Change()
 tipo_civil
End Sub

Private Sub cobEstado_Click()
 tipo_civil
End Sub

Private Sub cobEstado_Scroll()
 tipo_civil
End Sub

Private Sub cobSexo_Change()
 tipo_sexo
End Sub

Private Sub cobSexo_Click()
 tipo_sexo
 cobEstado_Click
End Sub

Private Sub cobSexo_Scroll()
 tipo_sexo
End Sub

Public Sub editar(estado As Boolean)
 Dim x As Byte
 For x = 0 To 31
 txtnombre(x).Enabled = estado
 Next x
 cobSexo.Enabled = estado
 DTPicker1.Enabled = estado
 cobEstado.Enabled = estado
 MonthView1.Enabled = estado
 cmdpaises.Enabled = estado
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift _
As Integer, ByVal CallbackField As String, CallbackDate As Date)
 DTPicker1_Click
End Sub

Private Sub DTPicker1_Click()
 Dim x(2) As Integer
 txtnombre.Item(24).Text = DTPicker1.Value
 txtnombre.Item(25).Text = optenerEdad(DTPickerFechaHoy.Year, DTPicker1.Year)
End Sub

Private Sub DTPicker1_CloseUp()
 DTPicker1_Click
End Sub

Private Sub DTPicker1_MouseMove(Button As Integer, Shift _
As Integer, x As Single, y As Single)
 funcion_coloriar_controles Index, &HFFC0FF   '&HC0FFC0
End Sub

Private Sub Form_Load()
 MDIPrincipal.Enabled = False
 foncion_color
 DTPicker1_Click
 definir_sexo
 cargarEstados
 DTPickerFechaHoy.ShowWhatsThis
 txtnombre.Item(25).Text = ""
 tipo_civil
 funcion_visualizar_datos_enMemoria
 editar False
 cargar_datos
End Sub

Private Sub funcion_visualizar_datos_enMemoria()
 On Error GoTo nose
  With agenda
  txtnombre(0).Text = .nombre.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(1).Text = .nombrex.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(2).Text = .apellidom.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(3).Text = .apellidop.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(4).Text = .telefono0.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(5).Text = .telefono1.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  DTPicker2.Value = .hora.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(7).Text = .celular0.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(15).Text = .celular1.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(14).Text = .ci.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(13).Text = .direccion0.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(12).Text = .direccion1.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  MonthView1.Value = .fecharegistro.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(10).Text = .pais.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(9).Text = .departamento.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(8).Text = .ciudad.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(31).Text = .calle0.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(30).Text = .calle1.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(29).Text = .calle2.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(28).Text = .email0.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(27).Text = .email1.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(26).Text = .email2.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(25).Text = .edad.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(24).Text = .fn.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  DTPicker1.Value = CDate(.fn.Item(frmvisualizar.ListView1.SelectedItem.Index).Key)
  txtnombre(16).Text = .facebook0.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(17).Text = .facebook1.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(18).Text = .facebook2.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(19).Text = .tuiter0.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(20).Text = .tuiter1.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(21).Text = .tuiter2.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(22).Text = .nCasa.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  txtnombre(23).Text = .Ecivil.Item(frmvisualizar.ListView1.SelectedItem.Index).Key
  End With
nose:
End Sub

Private Sub definir_sexo()
 With cobSexo
 .Clear
 .AddItem UCase("Male - Masuclino")
 .AddItem UCase("Female - Femenino")
 .ListIndex = 0
 txtnombre.Item(31).Text = UCase(cobSexo.Text)
 End With
End Sub

Public Sub tipo_sexo()
 txtnombre.Item(31).Text = UCase(cobSexo.Text)
 cargarEstados
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Cancel = 1
 Select Case frmbusqueda.Visible
 Case (False)
 MDIPrincipal.Enabled = True
 End Select
 desoprimr_boton 6
 Select Case MsgBox("Quieres Salir de la Hoja de Modificación", _
 vbYesNo + vbInformation, nombre_programa)
 Case (vbYes)
 Cancel = 0
 Unload Me
 End Select
End Sub

Public Sub resolverError()
 If control(25) = True Then
 frmdatos.txtnombre(25).Text = ""
 ElseIf control(22) = True Then
 frmdatos.txtnombre(22).Text = ""
 ElseIf control(4) = True Then
 frmdatos.txtnombre(4).Text = ""
 ElseIf control(5) = True Then
 frmdatos.txtnombre(5).Text = ""
 ElseIf control(6) = True Then
 frmdatos.txtnombre(6).Text = ""
 ElseIf control(7) = True Then
 frmdatos.txtnombre(7).Text = ""
 ElseIf control(14) = True Then
 frmdatos.txtnombre(14).Text = ""
 ElseIf control(15) = True Then
 frmdatos.txtnombre(15).Text = ""
 End If
 Dim x As Byte
 For x = 0 To 30
 control(x) = False
 Next x
End Sub

Private Sub funcion_modificar_datos()
 With agenda
 .nombre.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(0).Text
 .nombrex.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(1).Text
 .apellidom.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(2).Text
 .apellidop.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(3).Text
 .telefono0.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(4).Text
 .telefono1.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(5).Text
 .hora.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = DTPicker2.Value 'txtnombre(6).Text
 .celular0.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(7).Text
 .celular1.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(15).Text
 .ci.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(14).Text
 .direccion0.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(13).Text
 .direccion1.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(12).Text
 .fecharegistro.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = MonthView1.Value
 .pais.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(10).Text
 .departamento.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(9).Text
 .ciudad.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(8).Text
 .calle0.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(31).Text
 .calle1.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(30).Text
 .calle2.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(29).Text
 .email0.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(28).Text
 .email1.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(27).Text
 .email2.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(26).Text
 .edad.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(25).Text
 .fn.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(24).Text
 .facebook0.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(16).Text
 .facebook1.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(17).Text
 .facebook2.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(18).Text
 .tuiter0.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(19).Text
 .tuiter1.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(20).Text
 .tuiter2.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(21).Text
 .nCasa.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(22).Text
 .Ecivil.Item(frmvisualizar.ListView1.SelectedItem.Index).Key = txtnombre(23).Text
 End With
 Unload Me
 frmvisualizar.cmdVisualizar_Click
 ModDatos.guardar_archivo
End Sub

Private Sub ModificarHojadeAgenda_Click()
 cmdModificar_Click 0
End Sub

Private Sub permitiredicion_Click()
 cmdModificar_Click 3
End Sub

Private Sub Timer1_Timer()
 lblhora.Caption = DTPicker2.Value
 lblfecha.Caption = devolverDiasconCeros(MonthView1.Day)
 lblmes.Caption = meses(MonthView1.Month)
 lblanio.Caption = MonthView1.Year
 lbldia.Caption = devolverDias(Weekday(MonthView1))
End Sub

Public Function meses(ByVal mes As Integer)
 Select Case mes
 Case (1): meses = "Enero - January "
 Case (2): meses = "Febrero - February "
 Case (3): meses = "Marzo - March "
 Case (4): meses = "Abril -April "
 Case (5): meses = "Mayo - May "
 Case (6): meses = "Junio - June "
 Case (7): meses = "Julio - July "
 Case (8): meses = "Agosto - August "
 Case (9): meses = "Septiembre - September "
 Case (10): meses = "Octubre - October "
 Case (11): meses = "Noviembre - Nobember "
 Case (12): meses = "Diciembre - December "
 End Select
End Function
    
Public Function devolverDias(ByVal dia As Integer)
 Dim dias(1 To 7) As String
 dias(1) = "Domingo - Sunday"
 dias(2) = "Lunes - Monday"
 dias(3) = "Martes - Tuesday"
 dias(4) = "Miércoles - Wednesday"
 dias(5) = "Jueves - Thursday"
 dias(6) = "Viernes - Friday"
 dias(7) = "Sábado - Saturday"
 devolverDias = dias(dia)
End Function

Private Sub txtnombre_Change(Index As Integer)
 funcion_coloriar_controles Index, &HFF&
 comparar_ci
End Sub

Private Sub txtnombre_Click(Index As Integer)
 comparar_ci
End Sub

Private Sub txtnombre_KeyPress(Index As Integer, KeyAscii As Integer)
 On Error GoTo nose
 comparar_ci
 control(Index) = True
 soloLetrasMayusculas Index, KeyAscii
 Select Case (Index)
 Case (25)
 soloAceptarNumeros Index, "se permiten solo numeros para la " & textos(22) & " letras y otros caracteres no.", KeyAscii
 Case (22)
 soloAceptarNumeros Index, "se permiten solo numeros para la " & textos(30) & " pero separar con un digito por ejemplo utilizando el 0, letras y otros caracteres no.", KeyAscii
 Case (4)
 soloAceptarNumeros Index, "se permiten solo numeros para el " & textos(4) & " letras y otros caracteres no.", KeyAscii
 Case (5)
 soloAceptarNumeros Index, "se permiten solo numeros para el " & textos(5) & " letras y otros caracteres no.", KeyAscii
 Case (6)
 soloAceptarNumeros Index, "se permiten solo numeros para el " & textos(6) & " letras y otros caracteres no.", KeyAscii
 Case (7)
 soloAceptarNumeros Index, "se permiten solo numeros para el " & textos(7) & " letras y otros caracteres no.", KeyAscii
 Case (14)
 soloAceptarNumeros Index, "se permiten solo numeros para el " & textos(9) & " letras y otros caracteres no.", KeyAscii
 Case (15)
 soloAceptarNumeros Index, "se permiten solo numeros para la " & textos(8) & " letras y otros caracteres no.", KeyAscii
 End Select
nose:
End Sub

Public Sub soloLetrasMayusculas(ByRef Index As Integer, ByRef KeyAscii _
As Integer)
 If Index <= 3 Or Index = 8 Or Index = 9 Or Index = 10 Or Index = 11 Or _
 Index = 12 Or Index = 13 Or Index = 30 Or Index = 29 Or Index = 28 Or _
 Index = 27 Or Index = 26 Or Index = 16 Or Index = 17 Or Index = 18 Or _
 Index = 19 Or Index = 20 Or Index = 21 Or Index = 23 Then
 KeyAscii = Asc(UCase(Chr(KeyAscii)))
 If KeyAscii < Asc("A") Or KeyAscii > Asc("Z") Then
 If KeyAscii <> Asc("Ñ") Then
 End If
 End If
 End If
End Sub

Private Sub comparar_ci()
 ModFunciones_Publicas.virtual_ci txtnombre(14).Text
End Sub

Private Sub txtnombre_MouseMove(Index As Integer, Button _
As Integer, Shift As Integer, x As Single, y As Single)
 funcion_coloriar_controles Index, &HC000C0
End Sub

Private Sub foncion_color()
 Dim i As Integer
 For i = 0 To txtnombre.Count - 1
 txtnombre(i).Font.Bold = True
 Next
End Sub

Private Sub agregarya()
 Select Case MsgBox(" ¿ Quiere Agendar Otro Contacto ? ", _
 vbYesNo + vbInformation, "Agendario")
 Case (vbNo)
 MDIPrincipal.Toolbar1.Buttons(2).Value = tbrUnpressed
 Unload frmdatos
 End Select
End Sub

Private Sub destruir_datos()
 Dim d As Integer
 For d = 0 To 31
 txtnombre(d).Text = ""
 Next
End Sub

Private Sub desplazar()
 picHoja.Top = VScroll1.Value * 900
End Sub

Private Sub VScroll1_Change()
 desplazar
End Sub

Private Sub VScroll1_Scroll()
 desplazar
End Sub

Private Sub cmdlinkdefacebook_Click()
 If txtnombre(18).Text = "" Then
 x = ShellExecute(Me.hwnd, "Open", _
 "http://www.martinsoft0.blogspot.com.uy/", &O0, &O0, 0)
 Else
 x = ShellExecute(Me.hwnd, "Open", txtnombre(18).Text, &O0, &O0, 0)
 End If
End Sub

Private Sub cmdlinkfacebook_Click()
 If txtnombre(17).Text = "" Then
 x = ShellExecute(Me.hwnd, "Open", _
 "http://www.martinsoft0.blogspot.com.uy/", &O0, &O0, 0)
 Else
 x = ShellExecute(Me.hwnd, "Open", txtnombre(17).Text, &O0, &O0, 0)
 End If
End Sub

Private Sub cmdlinktuiter_Click()
 If txtnombre(20).Text = "" Then
 x = ShellExecute(Me.hwnd, "Open", _
 "http://www.martinsoft0.blogspot.com.uy/", &O0, &O0, 0)
 Else
 x = ShellExecute(Me.hwnd, "Open", txtnombre(20).Text, &O0, &O0, 0)
 End If
End Sub

Private Sub cmdlinktuitter_Click()
 If txtnombre(21).Text = "" Then
 x = ShellExecute(Me.hwnd, "Open", _
 "http://www.martinsoft0.blogspot.com.uy/", &O0, &O0, 0)
 Else
 x = ShellExecute(Me.hwnd, "Open", txtnombre(21).Text, &O0, &O0, 0)
 End If
End Sub

Private Sub cargar_datos()
 Open "etiquetas.opxl" For Input As 1
 Dim etiqueta As String
 Dim etiq As Integer
 Do While Not EOF(1)
 For etiq = 0 To 31
 Line Input #1, etiqueta
 textos(etiq) = escriptar.funcion_desescriptar(etiqueta)
 Next
 Loop
 Close #1
 cargarTextos
End Sub

Private Sub cargarTextos()
 Dim x As Integer
 For x = 0 To 31
 Label1(x).Caption = ""
 Next
 Label1(0).Caption = textos(0) & ":"
 Label1(1).Caption = textos(1) & ":"
 Label1(2).Caption = textos(2) & ":"
 Label1(3).Caption = textos(3) & ":"
 Label1(4).Caption = textos(4) & ":"
 Label1(5).Caption = textos(5) & ":"
 Label1(7).Caption = textos(7) & ":"
 Label1(15).Caption = textos(8) & ":"
 Label1(14).Caption = textos(9) & ":"
 Label1(13).Caption = textos(10) & ":"
 Label1(12).Caption = textos(11) & ":"
 Label1(11).Caption = textos(12) & ":"
 Label1(10).Caption = textos(13) & ":"
 Label1(9).Caption = textos(14) & ":"
 Label1(8).Caption = textos(15) & ":"
 Label1(31).Caption = textos(16) & ":"
 Label1(30).Caption = textos(17) & ":"
 Label1(29).Caption = textos(18) & ":"
 Label1(28).Caption = textos(19) & ":"
 Label1(27).Caption = textos(20) & ":"
 Label1(26).Caption = textos(21) & ":"
 Label1(25).Caption = textos(22) & ":"
 Label1(24).Caption = textos(23) & ":"
 Label1(17).Caption = textos(24) & ":"
 Label1(16).Caption = textos(25) & ":"
 Label1(18).Caption = textos(26) & ":"
 Label1(19).Caption = textos(27) & ":"
 Label1(20).Caption = textos(28) & ":"
 Label1(21).Caption = textos(29) & ":"
 Label1(22).Caption = textos(30) & ":"
 Label1(23).Caption = textos(31) & ":"
 Label3.Caption = textos(6) & ":"
End Sub




