---
title: Chart.PlotVisibleOnly property (Project)
keywords: vbapj.chm131628
f1_keywords:
- vbapj.chm131628
ms.prod: project-server
ms.assetid: 0745cf62-2625-3f5f-3a33-97709cabba33
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.PlotVisibleOnly property (Project)
 **True** if only visible cells are plotted. **False** if both visible and hidden cells are plotted. Read/write **Boolean**.

## Syntax

_expression_.**PlotVisibleOnly**

_expression_ A variable that represents a **[Chart](Project.Chart.md)** object.


## Example

The following example causes Project to plot only visible cells in the chart.


```vb
Sub PlotVisible()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple scalar chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    Debug.Print chartShape.Chart.PlotVisibleOnly
    chartShape.Chart.PlotVisibleOnly = True
End Sub
```


## Property value

 **BOOL**


## See also


[Chart Object](Project.chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]