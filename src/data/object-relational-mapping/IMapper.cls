VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Description("Provides an interface to implement an object-relational mapper.")
'@Folder("VBALibrary.Data.ObjectRelationalMapping")
'@Interface

Option Explicit

Public Function DeleteAll(ByVal obj As IMappable) As Long
End Function

Public Function DeleteMultiple(ByVal col As Collection) As Long
End Function

Public Sub DeleteSingle(ByVal obj As IMappable)
End Sub

Public Function GetAll(ByVal obj As IMappable) As Collection
End Function

Public Function GetMultiple(ByVal obj As IMappable, ByVal strQuery As String) As Collection
End Function

Public Function GetMultipleByFilter(ByVal obj As IMappable, ByVal strFilterCriteria As String) As Collection
End Function

Public Function GetSingle(ByRef obj As IMappable, ByVal vntPrimaryKey As Variant) As Boolean
End Function

Public Sub InsertMultiple(ByRef col As Collection)
End Sub

Public Sub InsertSingle(ByRef obj As IMappable)
End Sub

Public Function ItemExists(ByVal obj As IMappable) As Boolean
End Function

Public Sub UpdateMultiple(ByVal col As Collection)
End Sub

Public Sub UpdateSingle(ByVal obj As IMappable)
End Sub

Public Sub UpsertMultiple(ByVal col As Collection)
End Sub

Public Sub UpsertSingle(ByVal obj As IMappable)
End Sub
