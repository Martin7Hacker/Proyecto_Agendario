VERSION 5.00
Begin VB.Form frmAplicarPaises 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pa�ses del Mundo"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7365
   Icon            =   "frmAplicarPaises.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1425
   ScaleWidth      =   7365
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton cmdaplicar 
      Caption         =   "&Aplicar"
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.ComboBox compais 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   7095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800080&
      BorderWidth     =   2
      X1              =   0
      X2              =   7320
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pa�ses del mundo con sus Capitales"
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
      TabIndex        =   0
      Top             =   120
      Width           =   3090
   End
End
Attribute VB_Name = "frmAplicarPaises"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'* Open Source
'* System Application Software
'* Programa Paises de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************
Private Sub cargar_Paises()
compais.AddItem "Afganist�n: Kabul."
compais.AddItem "Albania: Tirana."
compais.AddItem "Alemania: Berl�n."
compais.AddItem "Andorra: Andorra la Vieja."
compais.AddItem "Angola: Luanda."
compais.AddItem "Antigua y Barbuda: Saint John."
compais.AddItem "Arabia Saudita: Riad."
compais.AddItem "Argelia: Argel."
compais.AddItem "Argentina: Buenos Aires."
compais.AddItem "Armenia: Erev�n."
compais.AddItem "Australia: Canberra."
compais.AddItem "Austria: Viena."
compais.AddItem "Azerbaiy�n: Bak�."
compais.AddItem "Bahamas: Nas�u."
compais.AddItem "Banglad�s: Daca."
compais.AddItem "Barbados: Bridgetown."
compais.AddItem "Bar�in: Manama."
compais.AddItem "B�lgica: Bruselas."
compais.AddItem "Belice: Belmop�n."
compais.AddItem "Ben�n: Porto Novo."
compais.AddItem "Bielorrusia: Minsk."
compais.AddItem "Birmania: Naipyid�."
compais.AddItem "Bolivia: Sucre."
compais.AddItem "Bosnia y Herzegovina: Sarajevo."
compais.AddItem "Botsuana: Gaborone."
compais.AddItem "Brasil: Brasilia."
compais.AddItem "Brun�i: Bandar Seri Begawan."
compais.AddItem "Bulgaria: Sof�a."
compais.AddItem "Burkina Faso: Uagadug�."
compais.AddItem "Burundi: Bujumbura."
compais.AddItem "But�n: Timbu."
compais.AddItem "Cabo Verde: Praia."
compais.AddItem "Camboya: Nom Pen."
compais.AddItem "Camer�n: Yaund�."
compais.AddItem "Canad�: Ottawa."
compais.AddItem "Catar: Doha."
compais.AddItem "Chad: Yamena."
compais.AddItem "Chile: Santiago de Chile."
compais.AddItem "China: Pek�n."
compais.AddItem "Chipre: Nicosia."
compais.AddItem "Ciudad del Vaticano: Ciudad del Vaticano."
compais.AddItem "Colombia: Bogot�."
compais.AddItem "Comoras: Moroni."
compais.AddItem "Corea del Norte: Pionyang."
compais.AddItem "Corea del Sur: Se�l."
compais.AddItem "Costa de Marfil: Yamusukro."
compais.AddItem "Costa Rica: San Jos�."
compais.AddItem "Croacia: Zagreb."
compais.AddItem "Cuba: La Habana."
compais.AddItem "Dinamarca: Copenhague."
compais.AddItem "Dominica: Roseau."
compais.AddItem "Ecuador: Quito."
compais.AddItem "Egipto: El Cairo."
compais.AddItem "El Salvador: San Salvador."
compais.AddItem "Emiratos �rabes Unidos: Abu Dabi."
compais.AddItem "Eritrea: Asmara."
compais.AddItem "Eslovaquia: Bratislava."
compais.AddItem "Eslovenia: Liubliana."
compais.AddItem "Espa�a: Madrid."
compais.AddItem "Estados Unidos: Washington D. C."
compais.AddItem "Estonia: Tallin"
compais.AddItem "Etiop�a: Ad�s Abeba."
compais.AddItem "Filipinas: Manila."
compais.AddItem "Finlandia: Helsinki."
compais.AddItem "Fiyi: Suva."
compais.AddItem "Francia: Par�s."
compais.AddItem "Gab�n: Libreville."
compais.AddItem "Gambia: Banjul."
compais.AddItem "Georgia: Tiflis."
compais.AddItem "Ghana: Acra."
compais.AddItem "Granada: Saint George."
compais.AddItem "Grecia: Atenas."
compais.AddItem "Guatemala: Ciudad de Guatemala."
compais.AddItem "Guyana: Georgetown."
compais.AddItem "Guinea: Conakri."
compais.AddItem "Guinea-Bis�u: Bis�u."
compais.AddItem "Guinea Ecuatorial: Malabo."
compais.AddItem "Hait�: Puerto Pr�ncipe."
compais.AddItem "Honduras: Tegucigalpa."
compais.AddItem "Hungr�a: Budapest."
compais.AddItem "India: Nueva Delhi."
compais.AddItem "Indonesia: Yakarta."
compais.AddItem "Irak: Bagdad."
compais.AddItem "Ir�n: Teher�n."
compais.AddItem "Irlanda: Dubl�n."
compais.AddItem "Islandia: Reikiavik."
compais.AddItem "Islas Marshall: Majuro."
compais.AddItem "Islas Salom�n: Honiara."
compais.AddItem "Israel: Jerusal�n."
compais.AddItem "Italia: Roma."
compais.AddItem "Jamaica: Kingston."
compais.AddItem "Jap�n: Tokio."
compais.AddItem "Jordania: Am�n."
compais.AddItem "Kazajist�n: Astan�."
compais.AddItem "Kenia: Nairobi."
compais.AddItem "Kirguist�n: Biskek."
compais.AddItem "Kiribati: Tarawa."
compais.AddItem "Kuwait: Kuwait."
compais.AddItem "Laos: Vienti�n."
compais.AddItem "Lesoto: Maseru."
compais.AddItem "Letonia: Riga."
compais.AddItem "L�bano: Beirut."
compais.AddItem "Liberia: Monrovia."
compais.AddItem "Libia: Tr�poli."
compais.AddItem "Liechtenstein: Vaduz."
compais.AddItem "Lituania: Vilna."
compais.AddItem "Luxemburgo: Luxemburgo."
compais.AddItem "Madagascar: Antananarivo."
compais.AddItem "Malasia: Kuala Lumpur."
compais.AddItem "Malaui: Lilong�e."
compais.AddItem "Maldivas: Mal�."
compais.AddItem "Mal�: Bamako."
compais.AddItem "Malta: La Valeta."
compais.AddItem "Marruecos: Rabat."
compais.AddItem "Mauricio: Port Louis."
compais.AddItem "Mauritania: Nuakchot."
compais.AddItem "M�xico: M�xico D. F."
compais.AddItem "Micronesia: Palikir."
compais.AddItem "Moldavia: Chisin�u."
compais.AddItem "M�naco: M�naco."
compais.AddItem "Mongolia: Ul�n Bator."
compais.AddItem "Montenegro: Podgorica."
compais.AddItem "Mozambique: Maputo."
compais.AddItem "Namibia: Windhoek."
compais.AddItem "Nauru: Yaren."
compais.AddItem "Nepal: Katmand�."
compais.AddItem "Nicaragua: Managua."
compais.AddItem "N�ger: Niamey."
compais.AddItem "Nigeria: Abuya."
compais.AddItem "Noruega: Oslo."
compais.AddItem "Nueva Zelanda: Wellington."
compais.AddItem "Oman:Mascate."
compais.AddItem "Pa�ses Bajos: �msterdam."
compais.AddItem "Pakist�n: Islamabad."
compais.AddItem "Palaos: Ngerulmud."
compais.AddItem "Panam�: Panam�."
compais.AddItem "Pap�a Nueva Guinea: Port Moresby."
compais.AddItem "Paraguay: Asunci�n."
compais.AddItem "Per�: Lima."
compais.AddItem "Polonia: Varsovia."
compais.AddItem "Portugal: Lisboa."
compais.AddItem "Reino Unido: Londres."
compais.AddItem "Rep�blica Centroafricana: Bangui."
compais.AddItem "Rep�blica Checa: Praga."
compais.AddItem "Rep�blica de Macedonia: Skopie."
compais.AddItem "Rep�blica del Congo: Brazzaville."
compais.AddItem "Rep�blica Democr�tica del Congo: Kinsasa."
compais.AddItem "Rep�blica Dominicana: Santo Domingo."
compais.AddItem "Rep�blica Sudafricana: Pretoria."
compais.AddItem "Ruanda: Kigali."
compais.AddItem "Ruman�a: Bucarest."
compais.AddItem "Rusia: Mosc�."
compais.AddItem "Samoa: Apia."
compais.AddItem "San Crist�bal y Nieves: Basseterre.."
compais.AddItem "San Marino: San Marino."
compais.AddItem "San Vicente y las Granadinas:  Kingstown."
compais.AddItem "Santa Luc�a: Castries."
compais.AddItem "Santo Tom� y Pr�ncipe: Santo Tom�."
compais.AddItem "Senegal: Dakar."
compais.AddItem "Serbia: Belgrado."
compais.AddItem "Seychelles: Victoria."
compais.AddItem "Sierra Leona: Freetown."
compais.AddItem "Singapur: Singapur."
compais.AddItem "Siria: Damasco."
compais.AddItem "Somalia: Mogadiscio."
compais.AddItem "Sri Lanka: Sri Jayawardenapura Kotte."
compais.AddItem "Suazilandia: Mbabane."
compais.AddItem "Sud�n: Jartum."
compais.AddItem "Sud�n del Sur: Yuba."
compais.AddItem "Suecia: Estocolmo."
compais.AddItem "Suiza: Berna."
compais.AddItem "Surinam: Paramaribo."
compais.AddItem "Tailandia: Bangkok."
compais.AddItem "Tanzania: Dodoma."
compais.AddItem "Tayikist�n: Dusamb�."
compais.AddItem "Timor Oriental: Dili."
compais.AddItem "Togo: Lom�."
compais.AddItem "Tonga: Nukualofa."
compais.AddItem "Trinidad y Tobago: Puerto Espa�a."
compais.AddItem "T�nez: T�nez."
compais.AddItem "Turkmenist�n: Asjabad."
compais.AddItem "Turqu�a: Ankara."
compais.AddItem "Tuvalu: Funafuti."
compais.AddItem "Ucrania: Kiev."
compais.AddItem "Uganda: Kampala."
compais.AddItem "Uruguay: Montevideo."
compais.AddItem "Uzbekist�n: Taskent."
compais.AddItem "Vanuatu: Port Vila."
compais.AddItem "Venezuela: Caracas."
compais.AddItem "Vietnam: Han�i."
compais.AddItem "Yemen: San�."
compais.AddItem "Yibuti: Yibuti."
compais.AddItem "Zambia: Lusaka."
compais.AddItem "Zimbabue: Harare."
End Sub

Private Sub cmdaplicar_Click()
frmdatos.txtnombre(10).Text = UCase(compais.Text)
frmModificar.txtnombre(10).Text = UCase(compais.Text)
Unload Me
End Sub

Private Sub cmdcancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
compais.Clear
cargar_Paises
End Sub
