'
'Work adapted from Colin Legg
'http://colinlegg.wordpress.com/2013/06/26/count-distinct-or-unique-values-vba-udf/
'

Public Function countdistinct( _
    ByRef rngToCheck As Range, _
    Optional ByVal blnCaseSensitive As Boolean = True _
                            ) As Variant
 
    Static dicDistinct As Object
 
    Dim varValues As Variant, varValue As Variant
    Dim lngCount As Long, lngRow As Long, lngCol As Long
 
    On Error GoTo ErrorHandler
    
    Set rngToCheck = Intersect(rngToCheck.Worksheet.UsedRange, rngToCheck)
 
    If Not rngToCheck Is Nothing Then
 
        'assign cell value(s) into memory so they
        'are faster to work with
        varValues = rngToCheck.Value
 
        'if rngToCheck is more than 1 cell then
        'varValues will be a 2 dimensional array
        If IsArray(varValues) Then
 
            If dicDistinct Is Nothing Then
                Set dicDistinct = CreateObject("Scripting.Dictionary")
                dicDistinct.CompareMode = BinaryCompare
            Else
                dicDistinct.RemoveAll
            End If
 
            For lngRow = LBound(varValues, 1) To UBound(varValues, 1)
                For lngCol = LBound(varValues, 2) To UBound(varValues, 2)
 
                    varValue = varValues(lngRow, lngCol)
 
                    'ignore error values
                    If Not IsError(varValue) Then
 
                        'ignore blank cells
                        'including formulae which return ""
                        If LenB(varValue) > 0 Then
 
                            'if we have a string then let's allow for case sensitivity
                            If VarType(varValue) = vbString Then
                                If Not blnCaseSensitive Then
                                    varValue = UCase(varValue)
                                End If
                            End If
 
                            If Not dicDistinct.Exists(varValue) Then
                                dicDistinct.Add varValue, vbNullString
                            End If
 
                        End If
                    End If
                Next lngCol
            Next lngRow
 
            lngCount = dicDistinct.Count
        Else
            'ignore if cell contains an error or is blank
            If Not IsError(varValues) Then
                If LenB(varValues) > 0 Then
                    lngCount = 1
                End If
            End If
        End If
    End If
 
    countdistinct = lngCount
 
    Exit Function
 
ErrorHandler:
    countdistinct = CVErr(xlErrValue)
 
End Function

