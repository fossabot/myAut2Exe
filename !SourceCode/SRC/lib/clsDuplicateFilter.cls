VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDuplicateFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private itemStore As Collection

Public Sub Clear()
   Set itemStore = New Collection
End Sub

Public Function IsUnique(uniqueItemID$) As Boolean
Attribute IsUnique.VB_UserMemId = 0
   On Error Resume Next
   itemStore.Add "", uniqueItemID
   IsUnique = Err = 0
End Function

Private Sub Class_Initialize()
   Clear
End Sub
