Attribute VB_Name = "TestKvpS"
Option Explicit
Option Private Module
'@IgnoreModule
'@TestModule
'@Folder("Tests")

#If twinbasic Then
    'Do nothing
#Else
'@TestModule
Option Private Module
'@IgnoreModule
'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
    'This method runs before every test in the module..
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub
#End If

Public Sub KvpSTests()
    
#If twinbasic Then
    Debug.Print CurrentProcedureName ; vbTab, vbTab, vbTab,
#Else
    Debug.Print ErrEx.LiveCallstack.ProcedureName; vbTab, vbTab,
#End If

    'Test01_IsKvpS
    Test02_CountKeysItems
    Test03_Clear
    Test04_RemoveByKey
    Test05_RemoveByIndex
    Test06a_Exists
    Test06b_HoldsKey
    Test06c_LacksKey
    Test06d_HoldsItem
    Test06e_LacksItem
    
    Test07_IndexByKey
    Test08_GetItemByIndex
    Test09_LetItemByIndex
    Test10_SetItemByIndex
    Test11_GetItem
    Test12_LetItem
    Test13_SetItem
    Test14_DuplicateKeys
    Test15_UniqueKeys
    'Test16_Reverse  reversing a dictionary doewsn't make sense
    
    Debug.Print "Testing completed"

End Sub

'@TestMethod("KvpS")
Private Sub Test01_IsKvpS()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(True, True, True)
    Dim myKvpS As KvpS = KvpS.Deb
    Dim myResult(0 To 2)  As Boolean

    'Act:
    myResult(0) = VBA.IsObject(myKvpS)
    Debug.Print VBA.TypeName(myKvpS)
    myResult(1) = "KvpS" = VBA.TypeName(myKvpS)
    myResult(2) = "KvpS" = myKvpS.TypeName
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName

    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("KvpS")
Private Sub Test02_CountKeysItems()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
       'On Error GoTo TestFail
    
        'Arrange:
        Dim myExpectedCount As Long = 6
        Dim myExpectedKeys As Variant = Split("Hello World Its A Nice Day", " ")
        Dim myExpectedItems As Variant = Array(10, 20, 30, 40, 50, 60)
        Dim myKvpS As KvpS = KvpS.Deb
        
        With myKvpS
            .Add "Hello", 10
            .Add "World", 20
            .Add "Its", 30
            .Add "A", 40
            .Add "Nice", 50
            .Add "Day", 60
        End With
        
        Dim myResultKeys  As Variant
        Dim myResultItems As Variant
        Dim myResultCount As Long
        
        'Act:
        myResultCount = myKvpS.Count
        myResultKeys = myKvpS.Keys
        myResultItems = myKvpS.Items
       
        'Assert.Strict:
        AssertStrictAreEqual myExpectedCount, myResultCount, myProcedureName
        AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
        AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
        
        TestExit:
        Exit Sub
        TestFail:
        Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub
    
'@TestMethod("KvpS")
Private Sub Test03_Clear()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
       'On Error GoTo TestFail
    
        'Arrange:
        Dim myExpected As Long = 0
        
        
        Dim myKvpS As KvpS = KvpS.Deb
        
        With myKvpS
            .Add "Hello", 10
            .Add "World", 20
            .Add "Its", 30
            .Add "A", 40
            .Add "Nice", 50
            .Add "Day", 60
        End With
        
        Dim myResult As Long
        
        'Act:
        myKvpS.Clear
        myResult = myKvpS.Count
        
        
        'Assert.Strict:
        AssertStrictAreEqual myExpected, myResult, myProcedureName
        
        TestExit:
        Exit Sub
        TestFail:
        Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub
    
'@TestMethod("KvpS")
Private Sub Test04_RemoveByKey()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpectedCount As Long = 4
    Dim myExpectedKeys As Variant = Split("World A Nice Day", " ")
    Dim myExpectedItems As Variant = Array(20, 40, 50, 60)
    
    Dim myKvpS As KvpS = KvpS.Deb
    
    With myKvpS
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    Dim myResultKeys  As Variant
    Dim myResultItems As Variant
    Dim myResultCount As Long
    
    'Act:
    myKvpS.Remove "Hello"
    myKvpS.Remove "Its"
    myResultCount = myKvpS.Count
    myResultKeys = myKvpS.Keys
    myResultItems = myKvpS.Items
    
    'Assert.Strict:
    AssertStrictAreEqual myExpectedCount, myResultCount, myProcedureName
    AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
    AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("KvpS")
Private Sub Test05_RemoveByIndex()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpectedCount As Long = 4
    Dim myExpectedKeys As Variant = Split("World Its Nice Day", " ")
    Dim myExpectedItems As Variant = Array(20, 30, 50, 60)
    
    Dim myKvpS As KvpS = KvpS.Deb
    
    With myKvpS
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    Dim myResultKeys  As Variant
    Dim myResultItems As Variant
    Dim myResultCount As Long
    
    'Act:
    myKvpS.RemoveByIndex 0&
    myKvpS.RemoveByIndex 2&
    myResultCount = myKvpS.Count
    myResultKeys = myKvpS.Keys
    myResultItems = myKvpS.Items
    
    'Assert.Strict:
    AssertStrictAreEqual myExpectedCount, myResultCount, myProcedureName
    AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
    AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("KvpS")
Private Sub Test06a_Exists()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(True, True, False, False)
    
    Dim myKvpS As KvpS = KvpS.Deb
    
    With myKvpS
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    Dim myResult(0 To 3) As Variant
    
    'Act:
    myResult(0) = myKvpS.Exists("World")
    myResult(1) = myKvpS.Exists("Its")
    myResult(2) = myKvpS.Exists("Theree")
    myResult(3) = myKvpS.Exists(" Its")
    
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("KvpS")
Private Sub Test06b_HoldsKey()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(True, True, False, False)
    
    Dim myKvpS As KvpS = KvpS.Deb
    
    With myKvpS
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    Dim myResult(0 To 3) As Variant
    
    'Act:
    myResult(0) = myKvpS.HoldsKey("World")
    myResult(1) = myKvpS.HoldsKey("Its")
    myResult(2) = myKvpS.HoldsKey("Theree")
    myResult(3) = myKvpS.HoldsKey(" Its")
    
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("KvpS")
Private Sub Test06c_LacksKey()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(False, False, True, True)
    
    Dim myKvpS As KvpS = KvpS.Deb
    
    With myKvpS
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    Dim myResult(0 To 3) As Variant
    
    'Act:
    myResult(0) = myKvpS.LacksKey("World")
    myResult(1) = myKvpS.LacksKey("Its")
    myResult(2) = myKvpS.LacksKey("There")
    myResult(3) = myKvpS.LacksKey(" Its")
    
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub
    
    
'@TestMethod("KvpS")
Private Sub Test06d_HoldsItem()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(True, True, False, False)
    
    Dim myKvpS As KvpS = KvpS.Deb
    
    With myKvpS
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    Dim myResult(0 To 3) As Variant
    
    'Act:
    myResult(0) = myKvpS.HoldsItem(10)
    myResult(1) = myKvpS.HoldsItem(50)
    myResult(2) = myKvpS.HoldsItem(42)
    myResult(3) = myKvpS.HoldsItem(-1)
    
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("KvpS")
Private Sub Test06e_LacksItem()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(False, False, True, True)
    
    Dim myKvpS As KvpS = KvpS.Deb
    
    With myKvpS
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    Dim myResult(0 To 3) As Variant
    
    'Act:
    myResult(0) = myKvpS.LacksItem(10)
    myResult(1) = myKvpS.LacksItem(50)
    myResult(2) = myKvpS.LacksItem(42)
    myResult(3) = myKvpS.LacksItem(-1)
    
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("KvpS")
Private Sub Test07_IndexByKey()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(1, 2, -1, -1)
    
    Dim myKvpS As KvpS = KvpS.Deb
    
    With myKvpS
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    Dim myResult(0 To 3) As Variant
    
    'Act:
    myResult(0) = myKvpS.IndexByKey("World")
    myResult(1) = myKvpS.IndexByKey("Its")
    myResult(2) = myKvpS.IndexByKey("Theree")
    myResult(3) = myKvpS.IndexByKey(" Its")
    
    'Assert.Strict:
    AssertStrictSequenceEquals myExpected, myResult, myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("KvpS")
Private Sub Test08_GetItemByIndex()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpected As Variant = Array(10, 30, 50, Null)
    
    Dim myKvpS As KvpS = KvpS.Deb
    
    With myKvpS
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    Dim myResult(0 To 3) As Variant
    
    'Act:
    myResult(0) = myKvpS.ItemByIndex(0)
    myResult(1) = myKvpS.ItemByIndex(2)
    myResult(2) = myKvpS.ItemByIndex(4)
    ' currently cHashD errors when out of range
   ' myResult(3) = myKvpS.ItemByIndex(7)
    
    'Assert.Strict:
    'This format is required as the VBA spec states that Null is not equal to Null
    ' so we cannot use sequence comparing
    AssertStrictAreEqual myExpected(0), myResult(0), myProcedureName
    AssertStrictAreEqual myExpected(1), myResult(1), myProcedureName
    AssertStrictAreEqual myExpected(2), myResult(2), myProcedureName
   ' AssertStrictAreEqual IsNull(myExpected(3)), IsNull(myResult(3)), myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("KvpS")
Private Sub Test09_LetItemByIndex()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpectedCount As Long = 6
    Dim myExpectedItems As Variant = Array(10, 20, 30, 42, 50, 60)
    Dim myExpectedKeys As Variant = Split("Hello World Its A Nice Day", " ")
    
    Dim myKvpS As KvpS = KvpS.Deb
    
    With myKvpS
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    'Act:
    Dim myVal As Long = 42&
    myKvpS.ItemByIndex(3) = myVal
    Dim myResultCount As Long = myKvpS.Count
    Dim myResultKeys As Variant = myKvpS.Keys
    Dim myResultItems As Variant = myKvpS.Items
    
    'Assert.Strict:
    AssertStrictAreEqual myExpectedCount, myResultCount, myProcedureName
    AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
    AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("KvpS")
Private Sub Test10_SetItemByIndex()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpectedCount As Long = 6
    Dim myExpectedKeys As Variant = Split("Hello World Its A Nice Day", " ")
    
    Dim myKvpS As KvpS = KvpS.Deb
    
    With myKvpS
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    Dim myCollection As Collection = New Collection
    
    With myCollection
        .Add 10
        .Add 20
        .Add 30
        .Add 40
        .Add 50
        .Add 60
    End With
    
    'Act:
    Set myKvpS.ItemByIndex(3) = myCollection
    Dim myResultCount As Long = myKvpS.Count
    Dim myResultKeys As Variant = myKvpS.Keys
    
    
    'Assert.Strict:
    AssertStrictAreEqual myExpectedCount, myResultCount, myProcedureName
    AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
    'ToDO: revise class so that Item is not needed
    AssertStrictAreEqual 40, myKvpS.Item("A")(4), myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

Private Sub Test11_GetItem()

    #If twinbasic Then
        myProcedureName = CurrentProcedureName
        myComponentName = CurrentComponentName
    #Else
        myProcedureName = ErrEx.LiveCallstack.ProcedureName
        myComponentName = ErrEx.LiveCallstack.ModuleName
    #End If
    
       'On Error GoTo TestFail
    
        'Arrange:
        Dim myExpected As Long = 50
       
        Dim myKvpS As KvpS = KvpS.Deb
        
        With myKvpS
            .Add "Hello", 10
            .Add "World", 20
            .Add "Its", 30
            .Add "A", 40
            .Add "Nice", 50
            .Add "Day", 60
        End With
        
        
        'Act:
        Dim myResult As Long = myKvpS.Item("Nice")
       
        'Assert.Strict:
        'This format is required as the VBA spec states that Null is not equal to Null
        ' so we cannot use sequence comparing
        AssertStrictAreEqual myExpected, myResult, myProcedureName
       
        
        TestExit:
        Exit Sub
        TestFail:
        Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
        Resume TestExit
    End Sub

'@TestMethod("KvpS")
Private Sub Test12_LetItem()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpectedCount As Long = 6
    Dim myExpectedItems As Variant = Array(10, 20, 30, 42, 50, 60)
    Dim myExpectedKeys As Variant = Split("Hello World Its A Nice Day", " ")
    Dim myKvpS As KvpS = KvpS.Deb
    
    With myKvpS
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    
    'Act:
    myKvpS.Item("A") = 42
    Dim myResultCount As Long = myKvpS.Count
    Dim myResultKeys As Variant = myKvpS.Keys
    Dim myResultItems As Variant = myKvpS.Items
    'Assert.Strict:
    'This format is required as the VBA spec states that Null is not equal to Null
    ' so we cannot use sequence comparing
    AssertStrictAreEqual myExpectedCount, myResultCount, myProcedureName
    AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
    AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("KvpS")
Private Sub Test13_SetItem()

#If twinbasic Then
    myProcedureName = CurrentProcedureName
    myComponentName = CurrentComponentName
#Else
    myProcedureName = ErrEx.LiveCallstack.ProcedureName
    myComponentName = ErrEx.LiveCallstack.ModuleName
#End If

   'On Error GoTo TestFail

    'Arrange:
    Dim myExpectedCount As Long = 6
    Dim myExpectedKeys As Variant = Split("Hello World Its A Nice Day", " ")
    Dim myKvpS As KvpS = KvpS.Deb
    
    With myKvpS
        .Add "Hello", 10
        .Add "World", 20
        .Add "Its", 30
        .Add "A", 40
        .Add "Nice", 50
        .Add "Day", 60
    End With
    
    Dim myCollection As Collection = New Collection
    
    With myCollection
        .Add 10
        .Add 20
        .Add 30
        .Add 40
        .Add 50
        .Add 60
    End With
    'Act:
    Set myKvpS.Item("A") = myCollection
    Dim myResultCount As Long = myKvpS.Count
    Dim myResultKeys As Variant = myKvpS.Keys
   
    'Assert.Strict:
    'This format is required as the VBA spec states that Null is not equal to Null
    ' so we cannot use sequence comparing
    AssertStrictAreEqual myExpectedCount, myResultCount, myProcedureName
    AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
    AssertStrictAreEqual 40, myKvpS.Item("A")(4), myProcedureName
    
    TestExit:
    Exit Sub
    TestFail:
    Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
    Resume TestExit
End Sub

'@TestMethod("KvpS")  ' KvpS does not support init
Private Sub Test14_DuplicateKeys()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

'     'On Error GoTo TestFail

'     'Arrange:
'     Dim myExpectedCount As Long = 8
'     Dim myExpectedItems As Variant = Array(10, 20, 30, 100, 40, 50, 60, 1000)
'     Dim myExpectedKeys As Variant = Split("Hello World Its Hello A Nice Day Hello", " ")
   
'     ' This line fails
'     Dim myKvpS As KvpS = KvpS.Deb
'     myKvpS.ReInit(ipensureuniquekeys:=False)
    
'     ' These two lines work OK
'     'Dim myKvpS As KvpS = KvpS.Deb
'     'myKvpS.ReInit(ensureuniquekeys:=False)
'     With myKvpS
'         .Add "Hello", 10
'         .Add "World", 20
'         .Add "Its", 30
'         .Add "Hello", 100
'         .Add "A", 40
'         .Add "Nice", 50
'         .Add "Day", 60
'         .Add "Hello", 1000
'     End With
    
'     'Act:
'     Dim myResultCount As Long = myKvpS.Count
'     Dim myResultKeys As Variant = myKvpS.Keys
'     Dim myResultItems As Variant = myKvpS.Items
   
'     'Assert.Strict:
'     AssertStrictAreEqual myExpectedCount, myResultCount, myProcedureName
'     AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
'     AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    
'     TestExit:
'     Exit Sub
'     TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
End Sub

'@TestMethod("KvpS") ' KvpS does not support init
Private Sub Test15_UniqueKeys()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

'     ' with unique keys ensured adding a duplicate key causes an error
'     ' which we deliberatelately ignore
'     ' but the duplicate keys do not get added
'     ' ***On error resume next must be enabled for this test to pass***
'    On Error Resume Next

'     'Arrange:
'     Dim myExpectedCount As Long = 6
'     Dim myExpectedItems As Variant = Array(10, 20, 30, 40, 50, 60)
'     Dim myExpectedKeys As Variant = Split("Hello World Its A Nice Day", " ")
'     Dim myKvpS As KvpS = KvpS.Deb
   
'     myKvpS.ReInit ipEnsureUniqueKeys:=True
    
'     With myKvpS
'         .Add "Hello", 10
'         .Add "World", 20
'         .Add "Its", 30
'         .Add "Hello", 100
'         .Add "A", 40
'         .Add "Nice", 50
'         .Add "Day", 60
'         .Add "Hello", 1000
'     End With
'    'On Error GoTo TestFail
'     'Act:
'     Dim myResultCount As Long = myKvpS.Count
'     Dim myResultKeys As Variant = myKvpS.Keys
'     Dim myResultItems As Variant = myKvpS.Items
   
'     'Assert.Strict:
'     'This format is required as the VBA spec states that Null is not equal to Null
'     ' so we cannot use sequence comparing
'     AssertStrictAreEqual myExpectedCount, myResultCount, myProcedureName
'     AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
'     AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    
'     TestExit:
'     Exit Sub
'     TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
End Sub


'** Reversing a dictionary doesn't amke sense
' '@TestMethod("KvpS")
' Private Sub Test16_Reverse()

' #If twinbasic Then
'     myProcedureName = CurrentProcedureName
'     myComponentName = CurrentComponentName
' #Else
'     myProcedureName = ErrEx.LiveCallstack.ProcedureName
'     myComponentName = ErrEx.LiveCallstack.ModuleName
' #End If

  
'    'On Error Resume Next

'     'Arrange:
'     Dim myExpectedCount As Long = 6
'     Dim myExpectedItems As Variant = Array(60, 50, 40, 30, 20, 10)
'     Dim myExpectedKeys As Variant = Split("Day Nice A Its World Hello", " ")
'     Dim myKvpS As KvpS = KvpS.Deb
   
  
    
'     With myKvpS
'         .Add "Hello", 10
'         .Add "World", 20
'         .Add "Its", 30
'         .Add "A", 40
'         .Add "Nice", 50
'         .Add "Day", 60
        
'     End With
'    'On Error GoTo TestFail
'     'Act:
'     Dim myReversed As KvpS = myKvpS.Reverse
'     Dim myResultCount As Long = myReversed.Count
'     Dim myResultKeys As Variant = myReversed.Keys
'     Dim myResultItems As Variant = myReversed.Items
   
'     'Assert.Strict:
'     'This format is required as the VBA spec states that Null is not equal to Null
'     ' so we cannot use sequence comparing
'     AssertStrictAreEqual myExpectedCount, myResultCount, myProcedureName
'     AssertStrictSequenceEquals myExpectedKeys, myResultKeys, myProcedureName
'     AssertStrictSequenceEquals myExpectedItems, myResultItems, myProcedureName
    
'     TestExit:
'     Exit Sub
'     TestFail:
'     Debug.Print myComponentName, myProcedureName, " raised an error: #" & Err.Number & " - " & Err.Description
'     Resume TestExit
' End Sub