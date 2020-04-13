---
title: Application.AlignTableCellBottom method (Project)
keywords: vbapj.chm1523
f1_keywords:
- vbapj.chm1523
ms.prod: project-server
ms.assetid: 3eedfcb4-eb75-163f-6c3a-4dde97ddb110
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.AlignTableCellBottom method (Project)
Aligns text at the bottom of the cell, for selected cells in a report table.

## Syntax

_expression_. `AlignTableCellBottom`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Return value

 **Boolean**


## Example

In the following example, the **AlignTableCells** macro aligns the text for all tables in the specified report.


```vb
Sub TestAlignReportTables()
    Dim reportName As String
    Dim alignment As String   ' The value can be "top", "center", or "bottom".
    
    reportName = "Align Table Cells Report"
    alignment = "top"
    
    AlignTableCells reportName, alignment
End Sub

' Align the text for all tables in a specified report.
Sub AlignTableCells(reportName As String, alignment As String)
    Dim theReport As Report
    Dim shp As Shape
    
    Set theReport = ActiveProject.Reports(reportName)
    
    ' Activate the report. If the report is already active,
    ' ignore the run-time error 1004 from the Apply method.
    On Error Resume Next
    theReport.Apply
    On Error GoTo 0
    
    For Each shp In theReport.Shapes
        Debug.Print "Shape: " & shp.Type & ", " & shp.Name
        
        If shp.HasTable Then
            shp.Select
            
            Select Case alignment
                Case "top"
                    AlignTableCellTop
                Case "center"
                    AlignTableCellVerticalCenter
                Case "bottom"
                    AlignTableCellBottom
                Case Else
                    Debug.Print "AlignTableCells error: " & vbCrLf _
                        & "alignment must be top, center, or bottom."
                End Select
        End If
    Next shp
End Sub
```


## See also


[Application Object](Project.Application.md)



[Report Object](Project.report.md)
[AlignTableCellTop Method](Project.application.aligntablecelltop.md)
[AlignTableCellVerticalCenter Method](Project.application.aligntablecellverticalcenter.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]