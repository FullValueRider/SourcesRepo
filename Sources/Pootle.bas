

Sub ttest()
    Dim myParray As pArray = pArray(10, 20, 30, 40)
    Dim myOarr As oArr(Of Long) = oArr.deb(myParray)
    Dim mySeq As gSeq(Of String) = gSeq.Deb(oarr.deb(pArray.Deb(10, 20, 30, 40, 50, 60)))
    
    Debug.Print mySeq.Count
    
End Sub



