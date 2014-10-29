'Modified from UDF COUNTDISTINCT http://colinlegg.wordpress.com/2013/06/26/count-distinct-or-unique-values-vba-udf/
'09/15/2013
'Requires reference to Microsoft Scripting Runtime library (scrrrun.dll) IDE.
'UDF uses dictionary objects for count in order to counteract inefficiency of unique COUNTIF functions native to Excel.
'Added capacity to specify rows to COUNTDISTINCT.
'Parameters:
    'rngForCat: range of cells which is used to restrict the rows used for COUNTDISTINCT (e.g. organization names).
    'CatRngContain: Refers to rows, only when rngForCat is exactly equals to CatRngContain, countdistinct is executed

Public Function countdistinct_by_category( _
    ByRef rngToCheck As Range, _
    ByRef rngForCat As Range, _
    ByRef CatRngContain As Variant, _
    Optional ByVal blnCaseSensitive As Boolean = True _
                            ) As Variant
 
    Static dicDistinct As Object
 
    Dim varValues As Variant, varValue As Variant, varValuesCat As Variant
    
    Dim lngCount As Long, lngRow As Long, lngCol As Long
 
    On Error GoTo ErrorHandler

'Set rngToCheck = Intersect(rngToCheck.Worksheet.UsedRange, rngToCheck)
'Set rngToCheck = Intersect(rngForCat.Worksheet.UsedRange, rngForCat)
    
    If Not rngToCheck Is Nothing Then
 
        'assign cell value(s) into memory so they are faster to work with
        
        varValues = rngToCheck.Value
        varValuesCat = rngForCat.Value
        
        'Debug.Print "The value of variable varValuesCat11 is: " & varValuesCat(1,1)
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
                  
                 If varValuesCat(lngRow, lngCol) = CatRngContain Then
                  
                    varValue = varValues(lngRow, lngCol)
                    Debug.Print "varValue =" & varValue
                    Debug.Print "lngRow lngCol =" & lngRow & lngCol
 
 
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
 
    countdistinct_by_category = lngCount
 
    Exit Function
 
ErrorHandler:
    countdistinct_by_category = CVErr(xlErrValue)
 
End Function



