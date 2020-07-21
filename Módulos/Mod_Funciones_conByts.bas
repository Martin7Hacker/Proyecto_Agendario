Attribute VB_Name = "Mod_Funciones_conByts"
'***************************************************************************
'* Open Source
'* System Application Software
'* Módulo Mod_Funciones_conByts de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************
Public Declare Function ShellExecute Lib _
"shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters _
As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'genero las colecciones para los datos

Public Const escript_0 = 0  'escripta el archivo a  byte:  0
Public Const escript_1 = 1  'escripta el archivo a  byte:  1
Public Const escript_2 = 2  'escripta el archivo a  byte:  2
Public Const escript_3 = 3  'escripta el archivo a  byte:  3
Public Const escript_4 = 4  'escripta el archivo a  byte:  4
Public Const escript_5 = 5  'escripta el archivo a  byte:  5
Public Const escript_6 = 6  'escripta el archivo a  byte:  6
Public Const escript_7 = 7  'escripta el archivo a  byte:  7
Public Const escript_8 = 8  'escripta el archivo a  byte:  8
Public Const escript_9 = 9  'escripta el archivo a  byte:  9
Public Const escript_10 = 10  'escripta el archivo a  byte:  10
Public Const escript_11 = 11  'escripta el archivo a  byte:  11
Public Const escript_12 = 12  'escripta el archivo a  byte:  12
Public Const escript_13 = 13  'escripta el archivo a  byte:  13
Public Const escript_14 = 14  'escripta el archivo a  byte:  14
Public Const escript_15 = 15  'escripta el archivo a  byte:  15
Public Const escript_16 = 16  'escripta el archivo a  byte:  16
Public Const escript_17 = 17  'escripta el archivo a  byte:  17
Public Const escript_18 = 18  'escripta el archivo a  byte:  18
Public Const escript_19 = 19  'escripta el archivo a  byte:  19
Public Const escript_20 = 20  'escripta el archivo a  byte:  20
Public Const escript_21 = 21  'escripta el archivo a  byte:  21
Public Const escript_22 = 22  'escripta el archivo a  byte:  22
Public Const escript_23 = 23  'escripta el archivo a  byte:  23
Public Const escript_24 = 24  'escripta el archivo a  byte:  24
Public Const escript_25 = 25  'escripta el archivo a  byte:  25
Public Const escript_26 = 26  'escripta el archivo a  byte:  26
Public Const escript_27 = 27  'escripta el archivo a  byte:  27
Public Const escript_28 = 28  'escripta el archivo a  byte:  28
Public Const escript_29 = 29  'escripta el archivo a  byte:  29
Public Const escript_30 = 30  'escripta el archivo a  byte:  30
Public Const escript_31 = 31  'escripta el archivo a  byte:  31
Public Const escript_32 = 32  'escripta el archivo a  byte:  32
Public Const escript_33 = 33  'escripta el archivo a  byte:  33
Public Const escript_34 = 34  'escripta el archivo a  byte:  34
Public Const escript_35 = 35  'escripta el archivo a  byte:  35
Public Const escript_36 = 36  'escripta el archivo a  byte:  36
Public Const escript_37 = 37  'escripta el archivo a  byte:  37
Public Const escript_38 = 38  'escripta el archivo a  byte:  38
Public Const escript_39 = 39  'escripta el archivo a  byte:  39
Public Const escript_40 = 40  'escripta el archivo a  byte:  40
Public Const escript_41 = 41  'escripta el archivo a  byte:  41
Public Const escript_42 = 42  'escripta el archivo a  byte:  42
Public Const escript_43 = 43  'escripta el archivo a  byte:  43
Public Const escript_44 = 44  'escripta el archivo a  byte:  44
Public Const escript_45 = 45  'escripta el archivo a  byte:  45
Public Const escript_46 = 46  'escripta el archivo a  byte:  46
Public Const escript_47 = 47  'escripta el archivo a  byte:  47
Public Const escript_48 = 48  'escripta el archivo a  byte:  48
Public Const escript_49 = 49  'escripta el archivo a  byte:  49
Public Const escript_50 = 50  'escripta el archivo a  byte:  50
Public Const escript_51 = 51  'escripta el archivo a  byte:  51
Public Const escript_52 = 52  'escripta el archivo a  byte:  52
Public Const escript_53 = 53  'escripta el archivo a  byte:  53
Public Const escript_54 = 54  'escripta el archivo a  byte:  54
Public Const escript_55 = 55  'escripta el archivo a  byte:  55
Public Const escript_56 = 56  'escripta el archivo a  byte:  56
Public Const escript_57 = 57  'escripta el archivo a  byte:  57
Public Const escript_58 = 58  'escripta el archivo a  byte:  58
Public Const escript_59 = 59  'escripta el archivo a  byte:  59
Public Const escript_60 = 60  'escripta el archivo a  byte:  60
Public Const escript_61 = 61  'escripta el archivo a  byte:  61
Public Const escript_62 = 62  'escripta el archivo a  byte:  62
Public Const escript_63 = 63  'escripta el archivo a  byte:  63
Public Const escript_64 = 64  'escripta el archivo a  byte:  64
Public Const escript_65 = 65  'escripta el archivo a  byte:  65
Public Const escript_66 = 66  'escripta el archivo a  byte:  66
Public Const escript_67 = 67  'escripta el archivo a  byte:  67
Public Const escript_68 = 68  'escripta el archivo a  byte:  68
Public Const escript_69 = 69  'escripta el archivo a  byte:  69
Public Const escript_70 = 70  'escripta el archivo a  byte:  70
Public Const escript_71 = 71  'escripta el archivo a  byte:  71
Public Const escript_72 = 72  'escripta el archivo a  byte:  72
Public Const escript_73 = 73  'escripta el archivo a  byte:  73
Public Const escript_74 = 74  'escripta el archivo a  byte:  74
Public Const escript_75 = 75  'escripta el archivo a  byte:  75
Public Const escript_76 = 76  'escripta el archivo a  byte:  76
Public Const escript_77 = 77  'escripta el archivo a  byte:  77
Public Const escript_78 = 78  'escripta el archivo a  byte:  78
Public Const escript_79 = 79  'escripta el archivo a  byte:  79
Public Const escript_80 = 80  'escripta el archivo a  byte:  80
Public Const escript_81 = 81  'escripta el archivo a  byte:  81
Public Const escript_82 = 82  'escripta el archivo a  byte:  82
Public Const escript_83 = 83  'escripta el archivo a  byte:  83
Public Const escript_84 = 84  'escripta el archivo a  byte:  84
Public Const escript_85 = 85  'escripta el archivo a  byte:  85
Public Const escript_86 = 86  'escripta el archivo a  byte:  86
Public Const escript_87 = 87  'escripta el archivo a  byte:  87
Public Const escript_88 = 88  'escripta el archivo a  byte:  88
Public Const escript_89 = 89  'escripta el archivo a  byte:  89
Public Const escript_90 = 90  'escripta el archivo a  byte:  90
Public Const escript_91 = 91  'escripta el archivo a  byte:  91
Public Const escript_92 = 92  'escripta el archivo a  byte:  92
Public Const escript_93 = 93  'escripta el archivo a  byte:  93
Public Const escript_94 = 94  'escripta el archivo a  byte:  94
Public Const escript_95 = 95  'escripta el archivo a  byte:  95
Public Const escript_96 = 96  'escripta el archivo a  byte:  96
Public Const escript_97 = 97  'escripta el archivo a  byte:  97
Public Const escript_98 = 98  'escripta el archivo a  byte:  98
Public Const escript_99 = 99  'escripta el archivo a  byte:  99
Public Const escript_100 = 100  'escripta el archivo a  byte:  100
Public Const escript_101 = 101  'escripta el archivo a  byte:  101
Public Const escript_102 = 102  'escripta el archivo a  byte:  102
Public Const escript_103 = 103  'escripta el archivo a  byte:  103
Public Const escript_104 = 104  'escripta el archivo a  byte:  104
Public Const escript_105 = 105  'escripta el archivo a  byte:  105
Public Const escript_106 = 106  'escripta el archivo a  byte:  106
Public Const escript_107 = 107  'escripta el archivo a  byte:  107
Public Const escript_108 = 108  'escripta el archivo a  byte:  108
Public Const escript_109 = 109  'escripta el archivo a  byte:  109
Public Const escript_110 = 110  'escripta el archivo a  byte:  110
Public Const escript_111 = 111  'escripta el archivo a  byte:  111
Public Const escript_112 = 112  'escripta el archivo a  byte:  112
Public Const escript_113 = 113  'escripta el archivo a  byte:  113
Public Const escript_114 = 114  'escripta el archivo a  byte:  114
Public Const escript_115 = 115  'escripta el archivo a  byte:  115
Public Const escript_116 = 116  'escripta el archivo a  byte:  116
Public Const escript_117 = 117  'escripta el archivo a  byte:  117
Public Const escript_118 = 118  'escripta el archivo a  byte:  118
Public Const escript_119 = 119  'escripta el archivo a  byte:  119
Public Const escript_120 = 120  'escripta el archivo a  byte:  120
Public Const escript_121 = 121  'escripta el archivo a  byte:  121

Public Sub desoprimr_boton(ByVal Item As Byte)
 MDIPrincipal.Toolbar1.Buttons.Item(Item).Value = tbrUnpressed
End Sub

Public Sub oprimir_boton(ByVal Item As Byte)
 MDIPrincipal.Toolbar1.Buttons.Item(Item).Value = tbrPressed
End Sub

Public Function optenerEdad(ByVal anoActual As Date _
, ByVal anoNacimiento As Date) As Integer
 Dim x(2) As Integer
 x(0) = CInt(anoNacimiento)
 x(1) = CInt(anoActual)
 x(2) = x(1) - x(0)
 optenerEdad = x(2)
End Function

Public Sub soloAceptarNumeros(ByVal Index As Integer, ByVal mensaje _
As String, ByVal KeyAscii As Integer)
 If Index = 25 Or Index = 4 Or Index = 5 Or Index = 6 Or Index = 7 Or _
 Index = 15 Or Index = 14 Or 22 Then
 If KeyAscii >= 48 And KeyAscii <= 57 Or KeyAscii = 127 Or KeyAscii = 8 Then
 Else
 Beep
 frmdatos.resolverError
 frmmensaje.mensajeXD mensaje
 frmmensaje.Show 1
 End If
 End If
End Sub

Public Sub openAyuda(ByVal control As Form)
 Select Case MsgBox("los cursos del programa de Agendario se encuentran en el blogspot de Martinsoft Quieres aceder a los cursos de los diferentes programas de Martinsoft Opciónal", _
 vbInformation + vbYesNo, "Agendario v1.0")
 Case (vbYes)
 Dim x As String
 x = ShellExecute(control.hwnd, "Open", "http://adf.ly/1TJmhG", &O0, &O0, 0)
 End Select
End Sub

Public Sub consultartiempo(ByVal control As Form)
 Select Case MsgBox("¿Quieres consultar el estado del tiempo?", _
 vbInformation + vbYesNo, "Agendario v1.0")
 Case (vbYes)
 Dim x As String
 x = ShellExecute(control.hwnd, "Open", _
 "http://adf.ly/1TJlXi", &O0, &O0, 0)
 End Select
End Sub

Public Sub AbrirWeb(ByVal control As Form, ByVal web As String)
 Dim x As String
 x = ShellExecute(control.hwnd, "Open", web, &O0, &O0, 0)
End Sub
   
Function devolverDiasconCeros(ByVal dia As Byte) As String
 If dia <= 9 Then
 devolverDiasconCeros = "0" & dia
 ElseIf dia >= 10 Then
 devolverDiasconCeros = dia
 End If
End Function

