---
title: Chart.ApplyCustomType method (Project)
ms.prod: project-server
ms.assetid: 2bfe88c2-198e-a039-ace6-4ba362ce09d6
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.ApplyCustomType method (Project)
Applies a custom chart type to a chart.

## Syntax

_expression_. `ApplyCustomType` _(ChartType,_ _TypeName)_

_expression_ A variable that represents a **[Chart](Project.Chart.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ChartType_|Required|**Office.XlChartType**|The type of chart.|
| _TypeName_|Optional|**Variant**|The name of the chart type.|
| _ChartType_|Required|XLCHARTTYPE||
| _TypeName_|Optional|**Variant**||

## Return value

 **Nothing**


## Example

The following example changes the chart type to a clustered 3D bar chart.


```vb
Sub SetChartType()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple scalar chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    Debug.Print "Chart type before: " & chartShape.Chart.ChartType
    chartShape.Chart.ApplyCustomType (xl3DBarClustered)
    Debug.Print "Chart type after: " & chartShape.Chart.ChartType
End Sub
```


## See also


[Chart Object](Project.chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]