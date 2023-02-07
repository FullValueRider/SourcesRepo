
' ' Size of types in 16 bit words
' Public Const varByteLen As long = 2
' Public Const varIntegerLen As long = 2
' Public Const varLongLen As Long = 2
' Public Const varLongLongLen As Long = 4
' Public Const varSingleLen = 2
' Public Const varDoubleLen = 4
' Public Const varDecimalLen = 8
' Public Const varObjectLen = 2
' Public Const varCurrencyLen = 4
' Public Const varDateLen = 4

' Private Declare PtrSafe Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (ptrToDest As Any, ptrToSource As Any, ByVal Count As Long)


' Public Function GetDataAsIntArray(ByRef ipItem As Variant) As Integer()

'         ' Vartype cannot be used freely with objects because if an object has a default member then
'         ' vartype will return the type of the default member rather than vbObject
'         ' Hence mymyVarType id set to vbObject and then overwritten if ipItem is not an object
'         Dim mymyVarType As VbVarType = IIf(VBA.IsObject(ipItem), vbObject, VBA.VarType(ipItem))
      
'         ' The data section of a variant starts at offset 8
'         Dim mySourcePtr As LongPtr = VarPtr(ipItem) + 8
        
'         ' myLen is the number of 16 bit words that the data of a specifi typoe will occupy
'         Dim myLen As Long
        
'         Select Case myVT
'             Case vbString
'                 ' The data part of a variant for a string is always a reference so we use strptr 
'                 ' to get the pointer to the actual data, which is held as ascw (16 bit words)
'                 ' and overwrite the existing value
'                 myLen = VBA.Len(ipItem)
'                 mySourcePtr = VBA.StrPtr(ipItem)
            
'             Case vbObject
'                 myLen = varObjectLen
'                 mySourcePtr = VBA.ObjPtr(ipItem)
                
'             Case vbCurrency
'                 myLen = varCurrencyLen
            
'             Case vbLong, vbInteger, vbByte
'                 myLen = varLongLen
        
'             Case vbDouble
'                 myLen = varDoubleLen
                
'             Case vbDate
'                 myLen = varDateLen
               
'             Case vbSingle
'                 myLen = varSingleLen
                
'             Case vbLongLong
'                 myLen = varLongLongLen
                
                
'             Case vbDecimal
'                 ' structure of a decimal is 
'                 'scale byte
'                 'sign byte
'                 'where scale/sign are a 16 bit word
'                 ' Hi long
'                 ' where high is a single 32 bit word
'                 ' low Long '32bit word
'                 ' med long '32bit word
'                 ' where low,med comprise one 64 bit value
'                 ' which a 7 x 16 bit words
'                 ' this means that the value ,when represented as a set of intergers,
'                 ' can look odd because 12 bytes for the integer part
'                 ' run 12,11,10,9,4,3,2,1,8,7,6,5
'                 'Or As integers
'                 ' 5,6,1,2,3,4
'                 myLen = varDecimalLen
'                 'mySourcePtr =
                
'         End Select
        
'         Dim myArray() As Integer
'         ReDim myArray(0 To myLen - 1)
      
'         Dim myDestPtr As LongPtr = GetArrayInfo(myArray).SAUdt.pvData
        
'         ' if a variant has the VT_BYREF flag set then the value contained by the variant
'         ' is a pointer to the actual data, not the data itself.
        
'         ' copy the two bytes that hold the myVarType Type
'         Dim myInteger As Integer = 0
'         CopyMem myInteger, ByVal VarPtr(ipItem), 2
        
'         ' check if we need to dereference a pointer to data
'         ' but not if we have a string or object
'         If (Not mymyVarType = vbString) And (Not mymyVarType = vbObject) Then
'             If myInteger & VT_BYREF Then
                
'                 CopyMem mySourcePtr, ByVal mySourcePtr, 4
                
'             End If
'         End If
'         ' we cna now get theh actual data
'         CopyMem ByVal myDestPtr, ByVal mySourcePtr, myLen * 2
        
'         Return myArray
        
' End Function


' Sub ttest()
'     ' need to resolve how decimal is stored in memeory
'     Dim myItem As Decimal = -79228162514264337593543950335D
'     Dim myArray As Variant
'     myArray = GetDataAsIntArray(myItem)
'     Debug.Print Hex(myArray(0))
'     Debug.Print Hex(myArray(1))
'     Debug.Print Hex(myArray(2))
'     Debug.Print Hex(myArray(3))
'     Debug.Print Hex(myArray(4))
'     Debug.Print Hex(myArray(5))
'     Debug.Print Hex(myArray(6))
' End Sub


' Public Function FindIndex(ByRef Key As Variant, Optional ipHash As Long) As Long

'     Dim myVarType As VbVarType = IIf(VBA.IsObject(Key), vbObject, VBA.VarType(Key))
    
'     Dim HTUB As Long = s.HashTableSize - 1
'     Dim myResult As Long = ipHash
'     ipHash = HTUB 'init the HashValue (all bits to 1)
    
'     Dim myDataAsIntArr As Variant = GetDataAsIntArray(Key)
    
'     ipHash = HashIt(ipHash, myDataAsIntArr, s.comparemethod)
'     Dim myIndex As Long
'     Select Case myVarType
'         Case vbString
            
'             If s.CompareMethod = vbBinaryCompare Then
                
'                 If FindIndex = -1 Then
'                     Exit Function                                               'it's a "Hash-Only" Calculation	
'                 End If
                
'                 For myIndex = 0 To s.HashTable(ipHash).Count - 1
'                     If myIndex < DynTakeOver Then
'                         FindIndex = s.HashTable(ipHash).DataIdxsStat(myIndex)
'                     Else
'                         FindIndex = s.HashTable(ipHash).DataIdxsDyn(myIndex)
'                     End If
'                     If VarType(s.Keys(FindIndex)) = myVarType Then
'                         If Key = s.Keys(FindIndex) Then
'                             Exit Function
'                         End If
'                     End If
'                 Next
'             Else
                
'                 If FindIndex = -1 Then
'                     Exit Function                                               'it's a "Hash-Only" Calculation	
'                 End If
                
'                 For myIndex = 0 To s.HashTable(ipHash).Count - 1
'                     If myIndex < DynTakeOver Then
'                         FindIndex = s.HashTable(ipHash).DataIdxsStat(myIndex)
'                     Else
'                         FindIndex = s.HashTable(ipHash).DataIdxsDyn(myIndex)
'                     End If
'                     If VarType(s.Keys(FindIndex)) = myVarType Then
'                         If StrComp(Key, s.Keys(FindIndex), s.CompareMethod) = 0 Then
'                             Exit Function
'                         End If
'                     End If
'                 Next
'             End If
        
'         Case vbObject
            
'             If FindIndex = -1 Then
'                 Exit Function                                                       'it's a "Hash-Only" Calculation	
'             End If
        
'             For myIndex = 0 To s.HashTable(ipHash).Count - 1
'                 If myIndex < DynTakeOver Then
'                     FindIndex = s.HashTable(ipHash).DataIdxsStat(myIndex)
'                 Else
'                     FindIndex = s.HashTable(ipHash).DataIdxsDyn(myIndex)
'                 End If
                
'                 If VarType(s.Keys(FindIndex)) = myVarType Then
'                     If Key Is s.Keys(FindIndex) Then
'                         Exit Function
'                     End If
'                 End If
'             Next
        
'         Case vbCurrency
            
'             If FindIndex = -1 Then
'                 Exit Function                                                           'it's a "Hash-Only" Calculation	
'             End If
            
'             For myIndex = 0 To s.HashTable(ipHash).Count - 1
'                 If myIndex < DynTakeOver Then
'                     FindIndex = s.HashTable(ipHash).DataIdxsStat(myIndex)
'                 Else
'                     FindIndex = s.HashTable(ipHash).DataIdxsDyn(myIndex)
'                 End If
'                 If VarType(s.Keys(FindIndex)) = myVarType Then
'                     If C = s.Keys(FindIndex) Then
'                         Exit Function
'                     End If
'                 End If
'             Next
        
'         Case vbLong, vbInteger, vbByte
            
            
'             If FindIndex = -1 Then
'                 Exit Function                                                           'it's a "Hash-Only" Calculation	
'             End If
            
'             For myIndex = 0 To s.HashTable(ipHash).Count - 1
'                 If myIndex < DynTakeOver Then
'                     FindIndex = s.HashTable(ipHash).DataIdxsStat(myIndex)
'                 Else
'                     FindIndex = s.HashTable(ipHash).DataIdxsDyn(myIndex)
'                 End If
                
'                 Select Case VarType(s.Keys(FindIndex))
'                     Case vbLong, vbInteger, vbByte
'                         If Key = s.Keys(FindIndex) Then
'                             Exit Function
'                         End If
'                 End Select
'             Next
    
'         Case vbDouble
            
            
'             If FindIndex = -1 Then
'                 Exit Function                                                   'it's a "Hash-Only" Calculation	
'             End If
                
'             For myIndex = 0 To s.HashTable(ipHash).Count - 1
'                 If myIndex < DynTakeOver Then
'                     FindIndex = s.HashTable(ipHash).DataIdxsStat(myIndex)
'                 Else
'                     FindIndex = s.HashTable(ipHash).DataIdxsDyn(myIndex)
'                 End If
'                 If VarType(s.Keys(FindIndex)) = myVarType Then
'                     If Key = s.Keys(FindIndex) Then
'                         Exit Function
'                     End If
'                 End If
'             Next
            
'         Case vbDate
            
'             If FindIndex = -1 Then
'                 Exit Function                                                   'it's a "Hash-Only" Calculation	
'             End If
                    
'             For myIndex = 0 To s.HashTable(ipHash).Count - 1
'                 If myIndex < DynTakeOver Then
'                     FindIndex = s.HashTable(ipHash).DataIdxsStat(myIndex)
'                 Else
'                     FindIndex = s.HashTable(ipHash).DataIdxsDyn(myIndex)
'                 End If
'                 If VarType(s.Keys(FindIndex)) = myVarType Then
'                     If Key = s.Keys(FindIndex) Then
'                         Exit Function
'                     End If
'                 End If
'             Next
        
'         Case vbSingle
            
            
'             If FindIndex = -1 Then
'                 Exit Function                                        'it's a "Hash-Only" Calculation	
'             End If
                
'             For myIndex = 0 To s.HashTable(ipHash).Count - 1
'                 If myIndex < DynTakeOver Then
'                     FindIndex = s.HashTable(ipHash).DataIdxsStat(myIndex)
'                 Else
'                     FindIndex = s.HashTable(ipHash).DataIdxsDyn(myIndex)
'                 End If
                
'                 If VarType(s.Keys(FindIndex)) = myVarType Then
'                     If Key = s.Keys(FindIndex) Then
'                         Exit Function
'                     End If
'                 End If
'             Next
            
'     End Select
    
'     FindIndex = -1
    

' End Function

' Public Function HashIt(ByRef ipHash As Long, ipDataAsIntArray As Variant, Optional ipCOmparemethod As VbCompareMethod = vbtextcompare) As Long

'     Dim myIndex As Long
'     Dim myHash As Long = ipHash
'     If ipCOmparemethod = vbBinaryCompare Then
    
'         For myIndex = LBound(ipDataAsIntArray) To UBound(ipDataAsIntArray)
'             myHash = (ipHash + ipDataAsIntArray(myIndex)) * HMul And HTUB
'         Next
        
'     Else
    
'         For myIndex = LBound(ipDataAsIntArray) To UBound(ipDataAsIntArray)
'             myHash = (ipHash + LWC(ipDataAsIntArray(myIndex))) * HMul And HTUB
'         Next
    	
'     End If
    
'     Return ipHash
    
' End Function