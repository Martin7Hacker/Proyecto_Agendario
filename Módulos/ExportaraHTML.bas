Attribute VB_Name = "ModExportaraHTML"
'***************************************************************************
'* Open Source
'* System Application Software
'* Módulo ModExportaraHTML de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************

Public Sub ExportarHTML_Chrome(ByVal arcnivoExterno _
As String, ByVal ListView As ListView, ByVal TEXTO_ENCABEZADO _
As String, ByVal TEXTO_PIE As String)
 Dim salto_linea, aleatorio_color, colorR As Boolean
 Dim codigo_Html As String
 Dim Fila, Columna, archivo As Integer
 aleatorio_color = True
 Screen.MousePointer = vbHourglass
 archivo = FreeFile
 Open arcnivoExterno For Output As archivo
 codigo_Html = codigo_Html & "<!DOCTYPE" & "html" & "PUBLIC" & "-//W3C//DTD XHTML" & "1.0" & "Transitional//EN" & "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd" & ">" & vbCrLf
 codigo_Html = codigo_Html & "<!--####################################################################-->" & vbCrLf
 codigo_Html = codigo_Html & "<!--#     CÓDIGO HTML GENERADO CON AGENDARIO V1.0 MARTINSOFT SOFTWARE  #-->" & vbCrLf
 codigo_Html = codigo_Html & "<!--####################################################################-->" & vbCrLf
 codigo_Html = codigo_Html & "<!--tipo de generación de código por salto de linea-->" & vbCrLf
 codigo_Html = codigo_Html & "<html " & " xmlns=" & "http://www.w3.org/1999/xhtml" & ">" & vbCrLf
 codigo_Html = codigo_Html & "<head>"
 codigo_Html = codigo_Html & "<meta http-equiv=" & "Content-Type" & "content=" & "text/html;charset=iso-8859-1" & "/>" & vbCrLf
 codigo_Html = codigo_Html & "<title>Exportaxión de Agendario v1.0</title>" & vbCrLf
 codigo_Html = codigo_Html & "<style " & "TYPE=" & "Text/css" & ">" & vbCrLf
 codigo_Html = codigo_Html & ".datos{font-style: inherit; font-weight: bold; color:#990077;};" & vbCrLf
 codigo_Html = codigo_Html & ".columna1 {color: #000000;background-color: #990066}" & vbCrLf
 codigo_Html = codigo_Html & ".Estilo4 {font-weight: bold; font-style: inherit;}" & vbCrLf
 codigo_Html = codigo_Html & ".columna{font-style: inherit; font-weight: bold; color: #000000; background-color: #990066; }" & vbCrLf
 codigo_Html = codigo_Html & "</style>" & vbCrLf
 codigo_Html = codigo_Html & "</head>" & vbCrLf
 codigo_Html = codigo_Html & "<body>" & vbCrLf
 codigo_Html = codigo_Html & "<p class=" & "Estilo3" & "> " & "<span " & "class=" & "Estilo4" & "> " & "Archivos de Agendario " & "v1.0 " & "<" & "/span> " & "<" & "/p>" & vbCrLf
 If con_hr1 = True Then
 codigo_Html = codigo_Html & "<hr align='Center'>" & vbCrLf
 End If
 codigo_Html = codigo_Html & "<table " & "width=5587" & " border=" & "1.5px" & " BorderColor=" & "#B1C3D9 " & "Class=" & "Columna" & ">" & vbCrLf
 codigo_Html = codigo_Html & "<tr>" & vbCrLf
 For repetir = 1 To ListView.ColumnHeaders.Count
 codigo_Html = codigo_Html & "<td " & "align=" & "center" & " class=" & "columna" & "> " & ListView.ColumnHeaders(repetir).Text & " </td>" & vbCrLf
 Next repetir
 For Fila = 1 To ListView.ListItems.Count
 codigo_Html = codigo_Html & "<tr>" & vbCrLf
 codigo_Html = codigo_Html & "<td align='center' class='columna1' style='text-decoration:underline; background-color: #9900CC;'>" & ListView.ListItems(Fila).Text & "</td>" & vbCrLf
 For Columna = 2 To ListView.ColumnHeaders.Count
 codigo_Html = codigo_Html & "<td align='center' class='columna1' style='text-decoration:underline; background-color: #9900CC;'>" & ListView.ListItems(Fila).SubItems(Columna - 1) & "</td>" & vbCrLf
 Next
 codigo_Html = codigo_Html & "</tr>" & vbCrLf
 Next
 If con_hr2 = True Then
 codigo_Html = codigo_Html & "<hr align='Center'>" & vbCrLf
 End If
 codigo_Html = codigo_Html & "</table>" & vbCrLf
 codigo_Html = codigo_Html & "<h5> Martinsoft Software </h5>" & vbCrLf
 codigo_Html = codigo_Html & "</body>" & vbCrLf
 codigo_Html = codigo_Html & "</html>" & vbCrLf
 Print #archivo, codigo_Html
 codigo_Html = ""
 Close
 Screen.MousePointer = vbNormal
 MsgBox " Archivo Html generado en: " & vbCrLf & arcnivoExterno, vbInformation
End Sub
