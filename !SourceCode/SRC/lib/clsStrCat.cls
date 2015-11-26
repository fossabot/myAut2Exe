VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStrCat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Buffer As String
Private BufferLen As Long
Private Pointer As Long

'Inhalt des internen Strings:
Public Property Let value(vData As String)
Attribute value.VB_UserMemId = 0

  Pointer = 1
  Concat vData
  
End Property
Public Property Get value() As String

  value = LeftB$(Buffer, Pointer - 1)
End Property

'Auf "Leerstring" setzen:
Public Sub Clear()
  Pointer = 1
End Sub

'String anhängen:
Public Sub Concat(ByRef value As String)
  Dim PointerNew As Long
  
  'Benötigten Buffer berechnen:
  PointerNew = Pointer + LenB(value)
  
  'Ggf. Buffer vergößern:
  If PointerNew > BufferLen Then
    Buffer = Buffer & Space$(PointerNew)
    BufferLen = LenB(Buffer)
  End If
  
  'String passend kopieren:
  MidB$(Buffer, Pointer) = value
  Pointer = PointerNew
End Sub

'String anhängen:
Public Sub ConcatVariant(ByRef value As Variant)
  Dim PointerNew As Long
  
  'Benötigten Buffer berechnen:
  PointerNew = Pointer + LenB(value)
  
  'Ggf. Buffer vergößern:
  If PointerNew > BufferLen Then
    Buffer = Buffer & Space$(PointerNew)
    BufferLen = LenB(Buffer)
  End If
  
  'String passend kopieren:
  MidB$(Buffer, Pointer) = value
  Pointer = PointerNew
End Sub



'Länge des internen Strings:
Public Property Get Length() As Long
  Length = Pointer \ 2
End Property





'Neues Objekt initialisieren:
Private Sub Class_Initialize()
  Clear
End Sub

