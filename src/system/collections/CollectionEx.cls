VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CollectionEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Description("Provides methods to extend the Collection class.")
'@Folder("VBALibrary.System.Collections")
'@PredeclaredId(True)
'@Remarks("This class has its hidden VB_PredeclaredId attribute set to true in order to create a static instance of the class without needing to be instantiated.")
Option Explicit

' =============================================================================

'@Description("Merges the values of two collections into a single new collection.")
'@Param("col1: A collection to be merged.")
'@Param("col2: A collection to be merged.")
'@Param("blnForceUnique: A value indicating whether the collection will discard duplicate values by forcing a unique key value.")
'@Returns("A collection object containing the merged objects of the two provided collections.")
Public Function MergeCollections(ByVal col1 As Collection, ByVal col2 As Collection, Optional ByVal blnForceUnique = False) As Collection

Dim colMergedCollection As Collection
Dim lngCollection1Count As Long
Dim lngCollection2Count As Long

    Set colMergedCollection = New Collection
    
    For lngCollection1Count = 1 To col1.Count
        colMergedCollection.Add col1.Item(lngCollection1Count)
    Next lngCollection1Count
    
    For lngCollection2Count = 1 To col2.Count
        If blnForceUnique Then
            If Not ItemExists(colMergedCollection, col2.Item(lngCollection2Count)) Then
                colMergedCollection.Add col2.Item(lngCollection2Count)
            End If
        Else
            colMergedCollection.Add col2.Item(lngCollection2Count)
        End If
    Next lngCollection2Count

    Set MergeCollections = colMergedCollection

End Function

' =============================================================================

'@Description("Determines whether the provided collection contains an item matching the provided key.")
'@Param("col: A collection containing items to be tested.")
'@Param("vntItem: The item to be tested.")
'@Returns("A value indicating whether the collection contains an item with the key.")
Public Function ItemExists(ByVal col As Collection, ByVal vntItem As Variant) As Boolean
    
Dim vntCurrentItem As Variant
Dim objCurrentItem As Object
Dim objItem As Object

    For Each vntCurrentItem In col
        If IsObject(vntCurrentItem) Then
            Set objCurrentItem = vntCurrentItem
            Set objItem = vntItem
            If objCurrentItem Is objItem Then
                ItemExists = True
                Exit Function
            End If
        Else
            If vntCurrentItem = vntItem Then
                ItemExists = True
                Exit Function
            End If
        End If
    Next vntCurrentItem
    ItemExists = False
    
End Function

' =============================================================================

'@Description("Determines whether the provided collection contains an item matching the provided key.")
'@Param("col: A collection containing items to be tested.")
'@Param("vntKey: The key to be tested.")
'@Returns("A value indicating whether the collection contains an item with the key.")
Public Function KeyExists(ByVal col As Collection, ByVal vntKey As Variant) As Boolean
    On Error Resume Next
    col.Item vntKey
    KeyExists = (Err.Number = 0)
    Err.Clear
End Function

' =============================================================================