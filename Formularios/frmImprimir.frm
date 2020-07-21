VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmImprimir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresion de Agendario"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6600
   Icon            =   "frmImprimir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6600
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   120
      TabIndex        =   8
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton cmdimprimirya 
      Caption         =   "&Imprimir ya..."
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin TabDlg.SSTab SSTab1 
         Height          =   2775
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4895
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         ForeColor       =   8388736
         TabCaption(0)   =   "Seleccionar impresora"
         TabPicture(0)   =   "frmImprimir.frx":0CCA
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "List1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "cmdselecionar"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdconfiguracionImpresion"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Información de la Impresora"
         TabPicture(1)   =   "frmImprimir.frx":19A4
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "List2"
         Tab(1).ControlCount=   1
         Begin VB.CommandButton cmdconfiguracionImpresion 
            Caption         =   "&Configuración de Impresión"
            Height          =   360
            Left            =   120
            TabIndex        =   7
            Top             =   2290
            Width           =   2895
         End
         Begin VB.ListBox List2 
            BackColor       =   &H8000000F&
            Height          =   2205
            Left            =   -74880
            TabIndex        =   6
            Top             =   360
            Width           =   5895
         End
         Begin VB.CommandButton cmdselecionar 
            Caption         =   "&Seleccionar "
            Height          =   360
            Left            =   3120
            TabIndex        =   5
            Top             =   2290
            Width           =   2895
         End
         Begin VB.ListBox List1 
            BackColor       =   &H8000000F&
            Height          =   1815
            Left            =   120
            TabIndex        =   4
            Top             =   390
            Width           =   5895
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Impresoras Disponibles"
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   1620
      End
   End
End
Attribute VB_Name = "frmImprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Open Source
'* System Application Software
'* Programa frmImprimir de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************
Option Explicit
Private Const NULLPTR = 0&: Private Const CCHDEVICENAME = 32
Private Const CCHFORMNAME = 32: Private Const DM_MODIFY = 8
Private Const DM_COPY = 2: Private Const DM_IN_BUFFER = DM_MODIFY
Private Const DM_OUT_BUFFER = DM_COPY: Private Const DMORIENT_PORTRAIT = 1
Private Const DMORIENT_LANDSCAPE = 2: Private Const DMRES_DRAFT = (-1)
Private Const DMRES_HIGH = (-4): Private Const DMRES_LOW = (-2)
Private Const DMRES_MEDIUM = (-3): Private Const DMTT_BITMAP = 1
Private Const DMTT_DOWNLOAD = 2: Private Const DMTT_DOWNLOAD_OUTLINE = 4
Private Const DMTT_SUBDEV = 3: Private Const DMCOLOR_COLOR = 2
Private Const DMCOLOR_MONOCHROME = 1
Private Type DEVMODE
 dmDeviceName(1 To CCHDEVICENAME) As Byte
 dmSpecVersion As Integer
 dmDriverVersion As Integer
 dmSize As Integer
 dmDriverExtra As Integer
 dmFields As Long
 dmOrientation As Integer
 dmPaperSize As Integer
 dmPaperLength As Integer
 dmPaperWidth As Integer
 dmScale As Integer
 dmCopies As Integer
 dmDefaultSource As Integer
 dmPrintQuality As Integer
 dmColor As Integer
 dmDuplex As Integer
 dmYResolution As Integer
 dmTTOption As Integer
 dmCollate As Integer
 dmFormName(1 To CCHFORMNAME) As Byte
 dmUnusedPadding As Integer
 dmBitsPerPel As Integer
 dmPelsWidth As Long
 dmPelsHeight As Long
 dmDisplayFlags As Long
 dmDisplayFrequency As Long
End Type
Private Declare Function OpenPrinter Lib _
"winspool.drv" Alias "OpenPrinterA" ( _
ByVal pPrinterName As String, _
phPrinter As Long, _
ByVal pDefault As Long) As Long
Private Declare Function DocumentProperties Lib _
"winspool.drv" Alias _
"DocumentPropertiesA" ( _
ByVal hwnd As Long, _
ByVal hPrinter As Long, _
ByVal pDeviceName As String, _
pDevModeOutput As Any, _
pDevModeInput As Any, _
ByVal fMode As Long) As Long
Private Declare Function ClosePrinter Lib "winspool.drv" ( _
ByVal hPrinter As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
hpvDest As Any, _
hpvSource As Any, _
ByVal cbCopy As Long)

Private Sub cmdcancelar_Click()
 Unload Me
End Sub

Private Sub cmdconfiguracionImpresion_Click()
 Dim printDlg As PrinterDlg: Set printDlg = New PrinterDlg
 printDlg.PrinterName = Printer.DeviceName
 printDlg.DriverName = Printer.DriverName
 printDlg.Port = Printer.Port
 printDlg.PaperBin = Printer.PaperBin
 printDlg.Flags = VBPrinterConstants.cdlPDNoSelection _
 Or VBPrinterConstants.cdlPDNoPageNums _
 Or VBPrinterConstants.cdlPDReturnDC
 Printer.TrackDefault = False
 If Not printDlg.ShowPrinter(Me.hwnd) Then
 Debug.Print "Cancel Selected"
 Exit Sub
 End If
 Dim NewPrinterName, strsetting As String: Dim objPrinter As Printer
 NewPrinterName = UCase$(printDlg.PrinterName)
 If Printer.DeviceName <> NewPrinterName Then
 For Each objPrinter In Printers
 If UCase$(objPrinter.DeviceName) = NewPrinterName Then
 Set Printer = objPrinter
 End If
 Next
 End If
 Printer.Copies = printDlg.Copies
 Printer.Orientation = printDlg.Orientation
 Printer.ColorMode = printDlg.ColorMode
 Printer.Duplex = printDlg.Duplex
 Printer.PaperBin = printDlg.PaperBin
 Printer.PaperSize = printDlg.PaperSize
 Printer.PrintQuality = printDlg.PrintQuality
 With Printer
 Debug.Print .DeviceName
 If .Orientation = 1 Then
 strsetting = "Portrait. "
 Else
 strsetting = "Landscape. "
 End If
 Debug.Print "Copies = " & .Copies, "Orientation = " & _
 strsetting
 If .ColorMode = 1 Then
 strsetting = "Black and White. "
 Else
 strsetting = "Color. "
 End If
 Debug.Print "ColorMode = " & strsetting
 If .Duplex = 1 Then
 strsetting = "None. "
 ElseIf .Duplex = 2 Then
 strsetting = "Horizontal/Long Edge. "
 ElseIf .Duplex = 3 Then
 strsetting = "Vertical/Short Edge. "
 Else
 strsetting = "Unknown. "
 End If
 Debug.Print "Duplex = " & strsetting
 Debug.Print "PaperBin = " & .PaperBin
 Debug.Print "PaperSize = " & .PaperSize
 Debug.Print "PrintQuality = " & .PrintQuality
 If (printDlg.Flags And VBPrinterConstants.cdlPDPrintToFile) = _
 VBPrinterConstants.cdlPDPrintToFile Then
 Debug.Print "Print to File Selected"
 Else
 Debug.Print "Print to File Not Selected"
 End If
 Debug.Print "hDC = " & printDlg.hdc
 End With
 Exit Sub
End Sub

Private Sub cmdimprimirya_Click()
 On Error GoTo nose
 Dim i As Integer, j As Integer
 For i = 1 To frmvisualizar.ListView1.ListItems.Count
 Printer.Print frmvisualizar.ListView1.ListItems(i).Text & vbTab;
 For j = 1 To frmvisualizar.ListView1.ListItems(i).ListSubItems.Count
 Printer.Print frmvisualizar.ListView1.ListItems(i).SubItems(j) & vbTab;
 Next j
 Printer.Print Chr(7)
 Next i
 Printer.EndDoc
nose:
End Sub

Private Function Establecer_Impresora(ByVal NamePrinter _
As String) As Boolean
 On Error GoTo errsub
 Dim obj_Impresora As Object
 Set obj_Impresora = CreateObject("WScript.Network")
 obj_Impresora.setdefaultprinter NamePrinter
 Set obj_Impresora = Nothing
 Establecer_Impresora = True
 MsgBox "La impresora se cambió correctamente", vbInformation
 Exit Function
errsub:
 If Err.Number = 0 Then Exit Function
 Establecer_Impresora = False
 MsgBox "error: " & Err.Number & Chr(13) & "Description: " & Err.Description
 On Error GoTo 0
End Function

Private Sub cmdselecionar_Click()
 If List1.Selected(List1.ListIndex) Then
 Establecer_Impresora List1
 End If
End Sub

Private Sub Form_Load()
 Mod_Funciones_conByts.oprimir_boton 13
 optener_configuracionImpresora
 Dim x As Printer, impr As String
 For Each x In Printers
 List1.AddItem x.DeviceName
 Next
 List1.ListIndex = 0
End Sub

Function StripNulls(OriginalStr As String) As String
 If (InStr(OriginalStr, Chr(0)) > 0) Then
 OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
 End If
 StripNulls = Trim(OriginalStr)
End Function

Function ByteToString(ByteArray() As Byte) As String
 Dim TempStr As String
 Dim i As Integer
 For i = 1 To CCHDEVICENAME
 TempStr = TempStr & Chr(ByteArray(i))
 Next i
 ByteToString = StripNulls(TempStr)
End Function

Function GetPrinterSettings(szPrinterName As String, _
 hdc As Long, ListBox As ListBox) As Boolean
 Dim hPrinter, nSize As Long: Dim pDevMode As DEVMODE
 Dim aDevMode() As Byte: Dim TempStr As String
 If OpenPrinter(szPrinterName, hPrinter, NULLPTR) Then
 nSize = DocumentProperties(NULLPTR, hPrinter, szPrinterName, NULLPTR, NULLPTR, 0)
 ReDim aDevMode(1 To nSize)
 nSize = DocumentProperties(NULLPTR, hPrinter, szPrinterName, _
 aDevMode(1), NULLPTR, DM_OUT_BUFFER)
 Call CopyMemory(pDevMode, aDevMode(1), Len(pDevMode))
 With ListBox
 .Clear ' Limpia el Listbox
 .AddItem "Printer Name: " & ByteToString(pDevMode.dmDeviceName)
 If pDevMode.dmOrientation = DMORIENT_PORTRAIT Then
 TempStr = "PORTRAIT"
 ElseIf pDevMode.dmOrientation = DMORIENT_LANDSCAPE Then
 TempStr = "LANDSCAPE"
 Else
 TempStr = "UNDEFINED"
 End If
 .AddItem "Orientación del papel: " & TempStr
 Select Case pDevMode.dmPrintQuality
 Case DMRES_DRAFT
 TempStr = "DRAFT"
 Case DMRES_HIGH
 TempStr = "HIGH"
 Case DMRES_LOW
 TempStr = "LOW"
 Case DMRES_MEDIUM
 TempStr = "MEDIUM"
 Case Else ' positive value
 TempStr = CStr(pDevMode.dmPrintQuality) & " dpi"
 End Select
 .AddItem "Calidad de impresión: " & TempStr
 Select Case pDevMode.dmTTOption
 ' default for dot-matrix printers
 Case DMTT_BITMAP
 TempStr = "TrueType fonts as graphics"
 ' default for HP printers that use PCL
 Case DMTT_DOWNLOAD
 TempStr = "Downloads TrueType fonts as soft fonts"
 Case DMTT_SUBDEV ' default for PostScript printers
 TempStr = "Substitute device fonts for TrueType fonts"
 Case Else
 TempStr = "UNDEFINED"
 End Select
 .AddItem "TrueType Option: " & TempStr
 If pDevMode.dmColor = DMCOLOR_MONOCHROME Then
 TempStr = "MONOCHROME"
 ElseIf pDevMode.dmColor = DMCOLOR_COLOR Then
 TempStr = "COLOR"
 Else
 TempStr = "UNDEFINED"
 End If
 .AddItem "Color or Monochrome: " & TempStr
 If pDevMode.dmScale = 0 Then
 TempStr = "NONE"
 Else
 TempStr = CStr(pDevMode.dmScale)
 End If
 .AddItem "Zoom: " & TempStr
 .AddItem "Resolución: " & pDevMode.dmYResolution & " dpi"
 .AddItem "Copias: " & CStr(pDevMode.dmCopies)
 End With
 Call ClosePrinter(hPrinter)
 GetPrinterSettings = True
 Else
 GetPrinterSettings = False
 End If
End Function

Private Sub optener_configuracionImpresora()
 Dim ret As Boolean
 ret = GetPrinterSettings(Printer.DeviceName, Printer.hdc, List2)
 If ret = False Then
 MsgBox "No se pudo obtener la configuración de la impresora"
 End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Mod_Funciones_conByts.desoprimr_boton 13
End Sub

