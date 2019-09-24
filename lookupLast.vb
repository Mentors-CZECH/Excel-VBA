'Option Explicit
Public rowPointer As Integer
Function VLOOKUP_LAST_NH(regionSheet As String, product As String)
    Dim productRowArray As Variant
    Dim productNameArray As Variant
    With Application
        .Volatile
    End With
    productNameArray = Array("A", "B", "C", "D", "E", "F")
    productRowArray = Array(37, 71, 105, 139, 173)
    itemIndex = Application.Match(product, productNameArray, False)
    
    VLOOKUP_LAST_NH = Sheets(regionSheet).Range("D" & productRowArray(itemIndex - 1)).End(xlDown).Value
End Function

Function VLOOKUP_LAST_RR(regionSheet As String, product As String)
    Dim productRowArray As Variant
    Dim productNameArray As Variant
    With Application
        .Volatile
    End With
    productNameArray = Array("A", "B", "C", "D", "E", "F")
    productRowArray = Array(37, 71, 105, 139, 173)
    itemIndex = Application.Match(product, productNameArray, False)
    rowPointer = Sheets(regionSheet).Range("D" & productRowArray(itemIndex - 1)).End(xlDown).Row
    
    VLOOKUP_LAST_RR = Sheets(regionSheet).Range("G" & rowPointer).Value
End Function

Function VLOOKUP_LAST_RR_PROC(regionSheet As String, product As String)
    Dim productRowArray As Variant
    Dim productNameArray As Variant
    With Application
        .Volatile
    End With
    productNameArray = Array("A", "B", "C", "D", "E", "F")
    productRowArray = Array(37, 71, 105, 139, 173)
    itemIndex = Application.Match(product, productNameArray, False)
    rowPointer = Sheets(regionSheet).Range("D" & productRowArray(itemIndex - 1)).End(xlDown).Row
    
    VLOOKUP_LAST_RR_PROC = Sheets(regionSheet).Range("H" & rowPointer).Value
End Function
