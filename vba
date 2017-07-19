Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    On Error GoTo Whoa

    Application.EnableEvents = False

    
    If Not Target.Cells.CountLarge > 1 Then
        If Not Intersect(Target, Rows(28)) Is Nothing Then
            Target.Offset(1, -4).Select
        End If
        
    End If
    
    If Not Target.Cells.CountLarge > 1 Then
        If Not Intersect(Target, Rows(53)) Is Nothing Then
            Target.Offset(-24, 2).Select
        End If
        
    End If
        
    If Not Target.Cells.CountLarge > 1 Then
        If Not Intersect(Target, Rows(78)) Is Nothing Then
            Target.Offset(-24, 2).Select
        End If
        
    End If
    
    If Not Target.Cells.CountLarge > 1 Then
        If Not Intersect(Target, Rows(103)) Is Nothing Then
            Target.Offset(-24, 2).Select
        End If
        
    End If
        
    If Not Target.Cells.CountLarge > 1 Then
        If Not Intersect(Target, Rows(128)) Is Nothing Then
            Target.Offset(-24, 2).Select
        End If

    End If
        
    If Not Target.Cells.CountLarge > 1 Then
        If Not Intersect(Target, Rows(153)) Is Nothing Then
            Target.Offset(-24, 2).Select
        End If
        
    End If
        
    If Not Target.Cells.CountLarge > 1 Then
        If Not Intersect(Target, Rows(178)) Is Nothing Then
            Target.Offset(-24, 2).Select
        End If
        
    End If
        
    If Not Target.Cells.CountLarge > 1 Then
        If Not Intersect(Target, Rows(203)) Is Nothing Then
            Target.Offset(-24, 2).Select
        End If
        
    End If
        
        ' Jumo to next catg \\.

    If Not Target.Cells.CountLarge > 1 Then
        If Not Intersect(Target, Range("F28")) Is Nothing Then
            Target.Offset(1, -4).Select
        End If

    End If

    If Not Target.Cells.CountLarge > 1 Then
        If Not Intersect(Target, Range("F53")) Is Nothing Then
            Target.Offset(1, -4).Select
        End If

    End If
    
    If Not Target.Cells.CountLarge > 1 Then
        If Not Intersect(Target, Range("F78")) Is Nothing Then
            Target.Offset(1, -4).Select
        End If

    End If
    
    If Not Target.Cells.CountLarge > 1 Then
        If Not Intersect(Target, Range("F103")) Is Nothing Then
            Target.Offset(1, -4).Select
        End If

    End If
    
    If Not Target.Cells.CountLarge > 1 Then
        If Not Intersect(Target, Range("F128")) Is Nothing Then
            Target.Offset(1, -4).Select
        End If

    End If
    
    If Not Target.Cells.CountLarge > 1 Then
        If Not Intersect(Target, Range("F153")) Is Nothing Then
            Target.Offset(1, -4).Select
        End If

    End If
    
    If Not Target.Cells.CountLarge > 1 Then
        If Not Intersect(Target, Range("F178")) Is Nothing Then
            Target.Offset(1, -4).Select
        End If

    End If
    
    If Not Target.Cells.CountLarge > 1 Then
        If Not Intersect(Target, Range("F28")) Is Nothing Then
            Target.Offset(1, -4).Select
        End If

    End If
Letscontinue:
    Application.EnableEvents = True
    Exit Sub
Whoa:
    MsgBox Err.Description
    Resume Letscontinue
End Sub




