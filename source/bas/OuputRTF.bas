Attribute VB_Name = "OuputRTF"
Dim aTableColor(255, 3) As Variant

Public Sub RTFOutput(oDoc As oPDF.cPDF, strRTF, FontName)

    Dim intX   As Integer
    Dim strAux As String
    
    'Cargo la tabla de colores
    Call GetTableColor(strRTF)
    
    'Elimino llave inicial y final
    strAux = Left(Mid(strRTF, 2), Len(strRTF) - 2)
    
    For intX = 1 To Len(strRTF)
        
    Next
    
    
End Sub

Public Function GetTableColor(strRTF As String) As String

    Dim strAux   As String
    Dim intStart As Integer
    Dim intEnd   As Integer
    Dim aColors  As Variant
    Dim aRGB     As Variant
    
    intStart = InStr(strRTF, "{\colortbl") + Len("{\colortbl")
    intEnd = InStr(intStart, strRTF, "}")
    
    strAux = TrimChar(Mid(strRTF, intStart, intEnd - intStart), ";")
    aColors = Split(strAux, ";")
    For intStart = LBound(aColors) To UBound(aColors)
        strAux = TrimChar(aColors(intStart), "\")
        aRGB = Split(strAux, "\")
        For intEnd = LBound(aRGB) To UBound(aRGB)
            aTableColor(intStart, intEnd) = lVal(aRGB(intEnd))
        Next
    Next
    
End Function

'Igual que Trim pero se le manda el caracter a eliminar
Private Function TrimChar(ByVal strString As String, ByVal strChar As String) As String

    strString = Trim(strString)
    If Left(strString, 1) = strChar Then
        strString = Mid(strString, 2)
    End If
    If Right(strString, 1) = strChar Then
        strString = Left(strString, Len(strString) - 1)
    End If
    
    TrimChar = strString

End Function

'Igual que val pero lee de derecha a izda.
Private Function lVal(ByVal strString As Variant) As Double

    Dim intX    As Integer
    Dim dblAux  As Double

    strString = CStr(strString)
    For intX = Len(strString) To 1 Step -1
        If IsNumeric(Mid(strString, intX)) Then
            dblAux = CDbl(Mid(strString, intX))
        Else
            Exit For
        End If
    Next
    
    lVal = dblAux
    
End Function
