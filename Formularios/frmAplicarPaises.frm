VERSION 5.00
Begin VB.Form frmAplicarPaises 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Países del Mundo"
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
      Caption         =   "Países del mundo con sus Capitales"
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
compais.AddItem "Afganistán: Kabul."
compais.AddItem "Albania: Tirana."
compais.AddItem "Alemania: Berlín."
compais.AddItem "Andorra: Andorra la Vieja."
compais.AddItem "Angola: Luanda."
compais.AddItem "Antigua y Barbuda: Saint John."
compais.AddItem "Arabia Saudita: Riad."
compais.AddItem "Argelia: Argel."
compais.AddItem "Argentina: Buenos Aires."
compais.AddItem "Armenia: Ereván."
compais.AddItem "Australia: Canberra."
compais.AddItem "Austria: Viena."
compais.AddItem "Azerbaiyán: Bakú."
compais.AddItem "Bahamas: Nasáu."
compais.AddItem "Bangladés: Daca."
compais.AddItem "Barbados: Bridgetown."
compais.AddItem "Baréin: Manama."
compais.AddItem "Bélgica: Bruselas."
compais.AddItem "Belice: Belmopán."
compais.AddItem "Benín: Porto Novo."
compais.AddItem "Bielorrusia: Minsk."
compais.AddItem "Birmania: Naipyidó."
compais.AddItem "Bolivia: Sucre."
compais.AddItem "Bosnia y Herzegovina: Sarajevo."
compais.AddItem "Botsuana: Gaborone."
compais.AddItem "Brasil: Brasilia."
compais.AddItem "Brunéi: Bandar Seri Begawan."
compais.AddItem "Bulgaria: Sofía."
compais.AddItem "Burkina Faso: Uagadugú."
compais.AddItem "Burundi: Bujumbura."
compais.AddItem "Bután: Timbu."
compais.AddItem "Cabo Verde: Praia."
compais.AddItem "Camboya: Nom Pen."
compais.AddItem "Camerún: Yaundé."
compais.AddItem "Canadá: Ottawa."
compais.AddItem "Catar: Doha."
compais.AddItem "Chad: Yamena."
compais.AddItem "Chile: Santiago de Chile."
compais.AddItem "China: Pekín."
compais.AddItem "Chipre: Nicosia."
compais.AddItem "Ciudad del Vaticano: Ciudad del Vaticano."
compais.AddItem "Colombia: Bogotá."
compais.AddItem "Comoras: Moroni."
compais.AddItem "Corea del Norte: Pionyang."
compais.AddItem "Corea del Sur: Seúl."
compais.AddItem "Costa de Marfil: Yamusukro."
compais.AddItem "Costa Rica: San José."
compais.AddItem "Croacia: Zagreb."
compais.AddItem "Cuba: La Habana."
compais.AddItem "Dinamarca: Copenhague."
compais.AddItem "Dominica: Roseau."
compais.AddItem "Ecuador: Quito."
compais.AddItem "Egipto: El Cairo."
compais.AddItem "El Salvador: San Salvador."
compais.AddItem "Emiratos Árabes Unidos: Abu Dabi."
compais.AddItem "Eritrea: Asmara."
compais.AddItem "Eslovaquia: Bratislava."
compais.AddItem "Eslovenia: Liubliana."
compais.AddItem "España: Madrid."
compais.AddItem "Estados Unidos: Washington D. C."
compais.AddItem "Estonia: Tallin"
compais.AddItem "Etiopía: Adís Abeba."
compais.AddItem "Filipinas: Manila."
compais.AddItem "Finlandia: Helsinki."
compais.AddItem "Fiyi: Suva."
compais.AddItem "Francia: París."
compais.AddItem "Gabón: Libreville."
compais.AddItem "Gambia: Banjul."
compais.AddItem "Georgia: Tiflis."
compais.AddItem "Ghana: Acra."
compais.AddItem "Granada: Saint George."
compais.AddItem "Grecia: Atenas."
compais.AddItem "Guatemala: Ciudad de Guatemala."
compais.AddItem "Guyana: Georgetown."
compais.AddItem "Guinea: Conakri."
compais.AddItem "Guinea-Bisáu: Bisáu."
compais.AddItem "Guinea Ecuatorial: Malabo."
compais.AddItem "Haití: Puerto Príncipe."
compais.AddItem "Honduras: Tegucigalpa."
compais.AddItem "Hungría: Budapest."
compais.AddItem "India: Nueva Delhi."
compais.AddItem "Indonesia: Yakarta."
compais.AddItem "Irak: Bagdad."
compais.AddItem "Irán: Teherán."
compais.AddItem "Irlanda: Dublín."
compais.AddItem "Islandia: Reikiavik."
compais.AddItem "Islas Marshall: Majuro."
compais.AddItem "Islas Salomón: Honiara."
compais.AddItem "Israel: Jerusalén."
compais.AddItem "Italia: Roma."
compais.AddItem "Jamaica: Kingston."
compais.AddItem "Japón: Tokio."
compais.AddItem "Jordania: Amán."
compais.AddItem "Kazajistán: Astaná."
compais.AddItem "Kenia: Nairobi."
compais.AddItem "Kirguistán: Biskek."
compais.AddItem "Kiribati: Tarawa."
compais.AddItem "Kuwait: Kuwait."
compais.AddItem "Laos: Vientián."
compais.AddItem "Lesoto: Maseru."
compais.AddItem "Letonia: Riga."
compais.AddItem "Líbano: Beirut."
compais.AddItem "Liberia: Monrovia."
compais.AddItem "Libia: Trípoli."
compais.AddItem "Liechtenstein: Vaduz."
compais.AddItem "Lituania: Vilna."
compais.AddItem "Luxemburgo: Luxemburgo."
compais.AddItem "Madagascar: Antananarivo."
compais.AddItem "Malasia: Kuala Lumpur."
compais.AddItem "Malaui: Lilongüe."
compais.AddItem "Maldivas: Malé."
compais.AddItem "Malí: Bamako."
compais.AddItem "Malta: La Valeta."
compais.AddItem "Marruecos: Rabat."
compais.AddItem "Mauricio: Port Louis."
compais.AddItem "Mauritania: Nuakchot."
compais.AddItem "México: México D. F."
compais.AddItem "Micronesia: Palikir."
compais.AddItem "Moldavia: Chisináu."
compais.AddItem "Mónaco: Mónaco."
compais.AddItem "Mongolia: Ulán Bator."
compais.AddItem "Montenegro: Podgorica."
compais.AddItem "Mozambique: Maputo."
compais.AddItem "Namibia: Windhoek."
compais.AddItem "Nauru: Yaren."
compais.AddItem "Nepal: Katmandú."
compais.AddItem "Nicaragua: Managua."
compais.AddItem "Níger: Niamey."
compais.AddItem "Nigeria: Abuya."
compais.AddItem "Noruega: Oslo."
compais.AddItem "Nueva Zelanda: Wellington."
compais.AddItem "Oman:Mascate."
compais.AddItem "Países Bajos: Ámsterdam."
compais.AddItem "Pakistán: Islamabad."
compais.AddItem "Palaos: Ngerulmud."
compais.AddItem "Panamá: Panamá."
compais.AddItem "Papúa Nueva Guinea: Port Moresby."
compais.AddItem "Paraguay: Asunción."
compais.AddItem "Perú: Lima."
compais.AddItem "Polonia: Varsovia."
compais.AddItem "Portugal: Lisboa."
compais.AddItem "Reino Unido: Londres."
compais.AddItem "República Centroafricana: Bangui."
compais.AddItem "República Checa: Praga."
compais.AddItem "República de Macedonia: Skopie."
compais.AddItem "República del Congo: Brazzaville."
compais.AddItem "República Democrática del Congo: Kinsasa."
compais.AddItem "República Dominicana: Santo Domingo."
compais.AddItem "República Sudafricana: Pretoria."
compais.AddItem "Ruanda: Kigali."
compais.AddItem "Rumanía: Bucarest."
compais.AddItem "Rusia: Moscú."
compais.AddItem "Samoa: Apia."
compais.AddItem "San Cristóbal y Nieves: Basseterre.."
compais.AddItem "San Marino: San Marino."
compais.AddItem "San Vicente y las Granadinas:  Kingstown."
compais.AddItem "Santa Lucía: Castries."
compais.AddItem "Santo Tomé y Príncipe: Santo Tomé."
compais.AddItem "Senegal: Dakar."
compais.AddItem "Serbia: Belgrado."
compais.AddItem "Seychelles: Victoria."
compais.AddItem "Sierra Leona: Freetown."
compais.AddItem "Singapur: Singapur."
compais.AddItem "Siria: Damasco."
compais.AddItem "Somalia: Mogadiscio."
compais.AddItem "Sri Lanka: Sri Jayawardenapura Kotte."
compais.AddItem "Suazilandia: Mbabane."
compais.AddItem "Sudán: Jartum."
compais.AddItem "Sudán del Sur: Yuba."
compais.AddItem "Suecia: Estocolmo."
compais.AddItem "Suiza: Berna."
compais.AddItem "Surinam: Paramaribo."
compais.AddItem "Tailandia: Bangkok."
compais.AddItem "Tanzania: Dodoma."
compais.AddItem "Tayikistán: Dusambé."
compais.AddItem "Timor Oriental: Dili."
compais.AddItem "Togo: Lomé."
compais.AddItem "Tonga: Nukualofa."
compais.AddItem "Trinidad y Tobago: Puerto España."
compais.AddItem "Túnez: Túnez."
compais.AddItem "Turkmenistán: Asjabad."
compais.AddItem "Turquía: Ankara."
compais.AddItem "Tuvalu: Funafuti."
compais.AddItem "Ucrania: Kiev."
compais.AddItem "Uganda: Kampala."
compais.AddItem "Uruguay: Montevideo."
compais.AddItem "Uzbekistán: Taskent."
compais.AddItem "Vanuatu: Port Vila."
compais.AddItem "Venezuela: Caracas."
compais.AddItem "Vietnam: Hanói."
compais.AddItem "Yemen: Saná."
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
