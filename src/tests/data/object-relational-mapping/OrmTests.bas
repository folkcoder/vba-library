Attribute VB_Name = "OrmTests"
'@Description("Tests for IMapper and IMappable implementations")
'@Folder("VBALibrary.Tests.Data.ObjectRelationalMapping")
'@TestModule

Option Explicit
Option Private Module

Private Assert As Object
Private Fakes As Object

Private mobjOrm As IMapper
Private mobjTestClass As IMappable


' =============================================================================
' INITIALIZE / CLEANUP
' =============================================================================

'@ModuleInitialize
Public Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
    Set mobjOrm = New OrmAdodb
    Set mobjTestClass = New IMappableMock
    mobjOrm.DeleteAll mobjTestClass
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    Set Assert = Nothing
    Set Fakes = Nothing
    Set mobjOrm = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
End Sub

'@TestCleanup
Public Sub TestCleanup()
    mobjOrm.DeleteAll mobjTestClass
End Sub

' =============================================================================

'@TestMethod
Public Sub DeleteAll_Test()
On Error GoTo TestFail

    'Arrange
    Dim lngRecordsDeleted As Long

    InsertTestRecord "DeleteAll_Test"
    InsertTestRecord "DeleteAll_Test"
    InsertTestRecord "DeleteAll_Test"

    'Act:
    lngRecordsDeleted = mobjOrm.DeleteAll(mobjTestClass)

    'Assert:
    Assert.IsTrue lngRecordsDeleted = 3

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    
End Sub

' =============================================================================

'@TestMethod
Public Sub DeleteMultiple_Test()
On Error GoTo TestFail

    'Arrange
    Dim colDeleted As Collection
    Dim lngRecordsDeleted As Long

    Set colDeleted = New Collection
    colDeleted.Add InsertTestRecord("DeleteMultiple_Test")
    colDeleted.Add InsertTestRecord("DeleteMultiple_Test")
    colDeleted.Add InsertTestRecord("DeleteMultiple_Test")
    colDeleted.Add InsertTestRecord("DeleteMultiple_Test")
    colDeleted.Add InsertTestRecord("DeleteMultiple_Test")

    'Act:
    lngRecordsDeleted = mobjOrm.DeleteMultiple(colDeleted)

    'Assert:
    Assert.IsTrue lngRecordsDeleted = colDeleted.Count

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    
End Sub

' =============================================================================
'@TestMethod
Public Sub DeleteSingle_Test()
On Error GoTo TestFail

    'Arrange
    Dim obj As IMappableMock
    Set obj = InsertTestRecord("DeleteSingle_Test")

    'Act:
    mobjOrm.DeleteSingle obj

    'Assert:
    Assert.IsFalse mobjOrm.ItemExists(obj)

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    
End Sub

' =============================================================================

'@TestMethod
Public Sub GetAll_Test()
On Error GoTo TestFail
   
    'Arrange:
    Dim col As Collection
    
    InsertTestRecord "GetAll_Test"
    InsertTestRecord "GetAll_Test"
    InsertTestRecord "GetAll_Test"
    InsertTestRecord "GetAll_Test"
    InsertTestRecord "GetAll_Test"
    InsertTestRecord "GetAll_Test"
    
    'Act:
    Set col = mobjOrm.GetAll(mobjTestClass)

    'Assert:
    Assert.IsTrue col.Count = 6

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description

End Sub

' =============================================================================

'@TestMethod
Public Sub GetMultiple_Test()
On Error GoTo TestFail
    
    'Arrange:
    Dim strQuery As String
    Dim col As Collection
    
    InsertTestRecord "Find_GetMultiple_Test"
    InsertTestRecord "Find_GetMultiple_Test"
    InsertTestRecord "Ignore_GetMultiple_Test"
        
    
    'Act:
    strQuery = "SELECT * FROM TestClass WHERE TestName = 'Find_GetMultiple_Test'"
    Set col = mobjOrm.GetMultiple(mobjTestClass, strQuery)

    'Assert:
    Assert.IsTrue (col.Count = 2)

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    
End Sub

' =============================================================================

'@TestMethod
Public Sub GetMultipleByFilter_Test()
On Error GoTo TestFail
    
    'Arrange:
    Dim strFilter As String
    Dim col As Collection
    
    InsertTestRecord "Find_GetMultipleByFilter_Test"
    InsertTestRecord "Find_GetMultipleByFilter_Test"
    InsertTestRecord "Ignore_GetMultipleByFilter_Test"
        
    
    'Act:
    strFilter = "TestName = 'Find_GetMultipleByFilter_Test'"
    Set col = mobjOrm.GetMultipleByFilter(mobjTestClass, strFilter)

    'Assert:
    Assert.IsTrue col.Count = 2

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
    
End Sub

' =============================================================================

'@TestMethod
Public Sub GetSingleByPrimaryKey_Test()
On Error GoTo TestFail
   
    'Arrange:
    Dim objInsert As IMappableMock
    Dim objRetrieve As IMappableMock
    Dim lngPrimaryKey As Long
    
    Set objInsert = InsertTestRecord("GetSingleByPrimaryKey_Test")
    lngPrimaryKey = objInsert.TestingID
    
    'Act:
    Set objRetrieve = New IMappableMock
    
    'Assert:
    Assert.IsTrue mobjOrm.GetSingle(objRetrieve, lngPrimaryKey)

TestExit:
    Exit Sub
    
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description

End Sub

' =============================================================================

'@TestMethod
Public Sub InsertMultiple_Test()
On Error GoTo TestFail
    
    'Arrange:
    Dim col As Collection
    Dim colTestResults As Collection
    Dim objTestResult As IMappableMock
    
    Set col = New Collection
    col.Add CreateTestRecord("InsertMultiple_Test")
    col.Add CreateTestRecord("InsertMultiple_Test")
    col.Add CreateTestRecord("InsertMultiple_Test")
    col.Add CreateTestRecord("InsertMultiple_Test")
    col.Add CreateTestRecord("InsertMultiple_Test")

    'Act:
    mobjOrm.InsertMultiple col

    'Assert:
    Set colTestResults = mobjOrm.GetMultipleByFilter(mobjTestClass, "TestName = 'InsertMultiple_Test'")
    Assert.IsTrue colTestResults.Count = col.Count
    
    Set objTestResult = colTestResults.Item(1)
    Assert.IsTrue objTestResult.TestingID <> 0

TestExit:
    Exit Sub

TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description

End Sub

' =============================================================================

'@TestMethod
Public Sub InsertSingle_Test()
On Error GoTo TestFail
    
    'Arrange:
    Dim obj As IMappableMock
    Dim colTestResults As Collection

    'Act:
    Set obj = InsertTestRecord("InsertSingle_Test")

    'Assert:
    Set colTestResults = mobjOrm.GetMultipleByFilter(mobjTestClass, "TestName = 'InsertSingle_Test'")
    Assert.IsTrue colTestResults.Count = 1
    Assert.IsTrue obj.TestingID <> 0

TestExit:
    Exit Sub

TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description

End Sub

' =============================================================================

'@TestMethod
Public Sub UpdateMultiple_Test()
On Error GoTo TestFail
    
    'Arrange:
    Dim objUpdate1 As IMappableMock
    Dim objUpdate2 As IMappableMock
    Dim objUpdate3 As IMappableMock
    
    Dim colObjsToUpdate As Collection
    Dim colUpdatedItems As Collection
    
    Dim objUpdatedItem As IMappableMock
    
    Set objUpdate1 = InsertTestRecord("UpdateMultiple_Test")
    Set objUpdate2 = InsertTestRecord("UpdateMultiple_Test")
    Set objUpdate3 = InsertTestRecord("UpdateMultiple_Test")
    
    objUpdate1.TestColor = "Salmon"
    objUpdate2.TestColor = "Puce"
    objUpdate3.TestColor = "Mauve"
    
    Set colObjsToUpdate = New Collection
    colObjsToUpdate.Add objUpdate1
    colObjsToUpdate.Add objUpdate2
    colObjsToUpdate.Add objUpdate3

    'Act:
    mobjOrm.UpdateMultiple colObjsToUpdate
    
    Set colUpdatedItems = mobjOrm.GetMultipleByFilter(mobjTestClass, "TestName = 'UpdateMultiple_Test'")

    'Assert:
    Set objUpdatedItem = colUpdatedItems.Item(1)
    Assert.IsTrue objUpdatedItem.TestColor = "Salmon"
    
    Set objUpdatedItem = colUpdatedItems.Item(2)
    Assert.IsTrue objUpdatedItem.TestColor = "Puce"
    
    Set objUpdatedItem = colUpdatedItems.Item(3)
    Assert.IsTrue objUpdatedItem.TestColor = "Mauve"

TestExit:
    Exit Sub

TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description

End Sub

' =============================================================================

'@TestMethod
Public Sub UpdateSingle_Test()
On Error GoTo TestFail
    
    'Arrange:
    Dim obj As IMappableMock
    Dim objTestUpdate As IMappableMock
        
    Set obj = InsertTestRecord("UpdateSingle_Test")
    
    'Act:
    obj.TestColor = "Maroon"
    mobjOrm.UpdateSingle obj
    
    Set objTestUpdate = New IMappableMock
    mobjOrm.GetSingle objTestUpdate, obj.TestingID

    'Assert:
    Assert.IsTrue objTestUpdate.TestColor = "Maroon"

TestExit:
    Exit Sub

TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description

End Sub

' =============================================================================

'@TestMethod
Public Sub UpsertMultipleAsUpdate_Test()
On Error GoTo TestFail

    'Arrange:
    Dim objUpdate1 As IMappableMock
    Dim objUpdate2 As IMappableMock
    Dim objUpdate3 As IMappableMock
    
    Dim colObjsToUpdate As Collection
    Dim colUpdatedItems As Collection
    
    Dim objUpdatedItem As IMappableMock
    
    Set objUpdate1 = InsertTestRecord("UpsertMultipleAsUpdate_Test")
    Set objUpdate2 = InsertTestRecord("UpsertMultipleAsUpdate_Test")
    Set objUpdate3 = InsertTestRecord("UpsertMultipleAsUpdate_Test")
    
    objUpdate1.TestColor = "Salmon"
    objUpdate2.TestColor = "Puce"
    objUpdate3.TestColor = "Mauve"
    
    Set colObjsToUpdate = New Collection
    colObjsToUpdate.Add objUpdate1
    colObjsToUpdate.Add objUpdate2
    colObjsToUpdate.Add objUpdate3

    'Act:
    mobjOrm.UpsertMultiple colObjsToUpdate
    
    Set colUpdatedItems = mobjOrm.GetMultipleByFilter(mobjTestClass, "TestName = 'UpsertMultipleAsUpdate_Test'")

    'Assert:
    Set objUpdatedItem = colUpdatedItems.Item(1)
    Assert.IsTrue objUpdatedItem.TestColor = "Salmon"
    
    Set objUpdatedItem = colUpdatedItems.Item(2)
    Assert.IsTrue objUpdatedItem.TestColor = "Puce"
    
    Set objUpdatedItem = colUpdatedItems.Item(3)
    Assert.IsTrue objUpdatedItem.TestColor = "Mauve"

TestExit:
    Exit Sub

TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description

End Sub

' =============================================================================

'@TestMethod
Public Sub UpsertSingleAsInsert_Test()
On Error GoTo TestFail
    
    'Arrange:
    Dim obj As IMappableMock
    Dim colTestResults As Collection

    'Act:
    Set obj = CreateTestRecord("UpsertSingleAsInsert_Test")
    mobjOrm.UpsertSingle obj

    'Assert:
    Set colTestResults = mobjOrm.GetMultipleByFilter(mobjTestClass, "TestName = 'UpsertSingleAsInsert_Test'")
    Assert.IsTrue colTestResults.Count = 1
    Assert.IsTrue obj.TestingID <> 0

TestExit:
    Exit Sub

TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description

End Sub

' =============================================================================

'@TestMethod
Public Sub UpsertSingleAsUpdate_Test()

    'Arrange:
    Dim obj As IMappableMock
    Dim objTestUpdate As IMappableMock
    
    Set obj = InsertTestRecord("UpdateSingle_Test")
    
    'Act:
    obj.TestColor = "Maroon"
    mobjOrm.UpsertSingle obj
    
    Set objTestUpdate = New IMappableMock
    mobjOrm.GetSingle objTestUpdate, obj.TestingID

    'Assert:
    Assert.IsTrue objTestUpdate.TestColor = "Maroon"

TestExit:
    Exit Sub

TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description

End Sub


' =============================================================================
' HELPER METHODS
' =============================================================================

Private Function CreateTestRecord(ByVal strTestName As String) As IMappableMock

Dim obj As IMappableMock

    Set obj = New IMappableMock
    With obj
        .TestColor = "Chartreuse"
        .TestName = strTestName
        .TestNumber = 12345
    End With
    Set CreateTestRecord = obj

End Function

' =============================================================================

Private Function InsertTestRecord(ByVal strTestName As String) As IMappableMock

Dim obj As IMappableMock

    Set obj = New IMappableMock
    With obj
        .TestColor = "Chartreuse"
        .TestName = strTestName
        .TestNumber = 12345
    End With
    mobjOrm.InsertSingle obj
    Set InsertTestRecord = obj

End Function
