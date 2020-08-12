Sub sortAsc(startRange As Range, endRange As Range, key As Range)
    Call Range(startRange, endRange).Sort(Key1:=key, Order1:=xlAscending)
End Sub
