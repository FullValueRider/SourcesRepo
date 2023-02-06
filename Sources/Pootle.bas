



Public Sub test()
 
    Dim mySeq As Seq = Seq.Deb(5)
    Set mySeq.Item(1) = Seq.Deb(1, 2, 3, 4, 5)
    Set mySeq.Item(2) = Seq.Deb(10, 20, 30, 40, 50)
    Set mySeq.Item(3) = Seq.Deb(11, 21, 31, 41, 51)
    Set mySeq.Item(4) = Seq.Deb(51, 61, 71, 81, 91)
    Set mySeq.Item(5) = Seq.Deb(5, 20, 15, 20, 25)
    
    Fmt.Dbg "{0}", mySeq
    Dim myTransposed As Seq = Transpose(mySeq)
    Fmt.Dbg "{0}", myTransposed
End Sub

 '@Description("Transpose a seq of seq")
    Public Function Transpose(ByRef ipSeq As Seq) As Seq
    	
        ' first create a seq of seq with empty values that matches
        ' the shape of transposed ipSeq
        Dim myResult As Seq = Seq.Deb
        Dim myRowIndex As Long
        For myRowIndex = ipSeq.First.firstIndex To ipSeq.First.lastindex
            myResult.AddItems Seq.Deb(ipSeq.LastIndex)
        Next
        
        
        Dim myRows As IterItems = IterItems(ipSeq)
        Do
            Dim myRow As Long = myRows.Key(0)
            
            Dim myCols As IterItems = IterItems(myRows.Item(0))
            Do
                
                Dim myCol As Long = myCols.Key(0)
            	myResult.Item(myCol).Item(myRow) = ipSeq.Item(myRow).Item(myCol)
                
            Loop While myCols.MoveNext
            
        Loop While myRows.MoveNext
        ' fmt.dbg is also called on the seq that becomes ipSeq before we enter this method
        ' for debugging purposes
        Debug.Print
        Fmt.Dbg "{0}", ipSeq     ' should print a copy of the existing fmt.dbg printout - yes
        Fmt.Dbg "{0}", myResult  ' should print transposed data - no - prints an empty seq.
        Return myResult
        
    End Function
