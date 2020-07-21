VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "datos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'* Open Source
'* System Application Software - Funcines virtuales de Búsqueda
'* clase datos de Agendario v1.0
'* By : Martin Grasso Castrillo - for all Proyect USA
'* Fb : https://www.facebook.com/hacker.martin0
'***************************************************************************
Private mCol As Collection

Public Function Add(Key As String, Optional sKey As String) As variable
 Dim objNewMember As variable
 Set objNewMember = New variable
 objNewMember.Key = Key
 If Len(sKey) = 0 Then
 mCol.Add objNewMember
 Else
 mCol.Add objNewMember, sKey
 End If
 Set Add = objNewMember
 Set objNewMember = Nothing
End Function

Public Property Get Item(vntIndexKey As Variant) As variable
 Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
 Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
 mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
 Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
 Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
 Set mCol = Nothing
End Sub

Public Sub Limpiar()
 If mCol.Count = 1 Then
 Else
 Dim x As Long
 For x = -mCol.Count To mCol.Count
 On Error GoTo no_se
 Dim ix As Long
 If mCol.Count = 1 Then
 For ix = 1 To mCol.Count
 mCol.Remove (ix)
 Next
 Else
 For ix = 1 To mCol.Count - 1
 mCol.Remove (ix)
 Next
 mCol.Remove (1)
 End If
 Next
 End If
no_se:
End Sub
