---
title: Chart.DataTable property (Project)
ms.prod: project-server
ms.assetid: 858ba41c-a96c-0c3d-0faf-dcfcc448c6f9
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.DataTable property (Project)
Gets an  **Office.IMsoDataTable** object that represents the chart data table. Read-only **IMsoDataTable**.

## Syntax

_expression_.**DataTable**

_expression_ A variable that represents a **[Chart](Project.Chart.md)** object.


## Remarks

To see the  **IMsoDataTable** object, right-click in the Object Browser, and then choose **Show Hidden Members**.


## Example

The following example adds a data table with an outline border to the chart on the active report.


```vb
Sub ShowDataTable()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple scalar chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    With chartShape.Chart
        .HasDataTable = True
        .DataTable.HasBorderOutline = True
    End With
End Sub
```


## Property value

 **IMSODATATABLE**


## See also


[Chart Object](Project.chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]