---
title: Chart.SaveChartTemplate method (Project)
ms.prod: project-server
ms.assetid: 496eb522-d758-ea4c-1cd9-4884c3b44189
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.SaveChartTemplate method (Project)
Saves a custom chart template to the list of available chart templates or to a file.

## Syntax

_expression_. `SaveChartTemplate` _(bstrFileName)_

_expression_ A variable that represents a **[Chart](Project.Chart.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _bstrFileName_|Required|**String**|The name of the chart template.|
| _bstrFileName_|Required|**String**||

## Return value

 **Nothing**


## Remarks

By default, the **SaveChartTemplate** method saves the active chart to the user's chart template directory (for example `C:\Users\username.DOMAIN\AppData\Roaming\Microsoft\Templates\Charts`). If a UNC file path or URL is specified, the chart is saved to the specified location.


## Example

The following example saves the chart template in the  `C:\Project\VBA\Samples\My chart template.crtx` file.


```vb
Sub SaveATemplate()
    Dim chartShape As Shape
    Dim reportName As String
    
    reportName = "Simple scalar chart"
    Set chartShape = ActiveProject.Reports(reportName).Shapes(1)
    
    chartShape.Chart.SaveChartTemplate "C:\Project\VBA\Samples\My chart template"
End Sub
```


## See also


[Chart Object](Project.chart.md)
[SetDefaultChart Method](Project.chart.setdefaultchart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]