

Sub ttest()
    Dim myParray As pArray = pArray(10, 20, 30, 40)
    Dim myOarr As oArr(Of Long) = oArr.Deb(myParray)    '.Deb(myParray)
    Dim mySeq As gSeq(Of Long) = gSeq.Deb(myOarr)
    
    Debug.Print mySeq.Count
    
End Sub



