Option Explicit

Sub joinByDelimiter()

    Dim strArray() As Variant
    Dim rangeSize, i As Double
    Dim delimiter As String
    Dim setRange, tgtRange As Object
    Dim MyRange, MyCell As Range
    
    Set setRange = Application.InputBox(prompt:="Zadejte rozsah, který chcete spojit (Pozue ve sloupci)", Type:=8)
    Set tgtRange = Application.InputBox(prompt:="Vyberte buňku do které se vloží řetězec", Type:=8)
    
    delimiter = Application.InputBox(prompt:="Zadejte delimiter")
    rangeSize = setRange.Rows.Count
    ReDim strArray(0 To rangeSize)
    For Each MyCell In setRange
        strArray(i) = MyCell.Value
        i = i + 1
    Next MyCell

    tgtRange.Value = Join(strArray, delimiter)
End Sub
