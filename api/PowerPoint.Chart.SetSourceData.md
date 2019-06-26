---
title: Chart.SetSourceData method (PowerPoint)
keywords: vbapp10.chm66949
f1_keywords:
- vbapp10.chm66949
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.SetSourceData
ms.assetid: f82dd278-318f-5706-3506-a0992df142ef
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.SetSourceData method (PowerPoint)

Sets the source data range for the chart.


## Syntax

_expression_. `SetSourceData`( `_Source_`, `_PlotBy_` )

_expression_ A variable that represents a **[Chart](PowerPoint.Chart.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Source_|Required|**String**|The address of the chart data range that contains the source data.|
| _PlotBy_|Optional|**Variant**|Specifies the way the data will be plotted. Can be either of the following  **[XlRowCol](PowerPoint.XlRowCol.md)** constants: **xlColumns** or **xlRows**.|

## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example sets the source data range for the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1)
    If .HasChart Then
        .Chart.SetSourceData _
            Source:="='Sheet1'!$A$1:$D$5", _
            PlotBy:=xlColumns
    End If
End With
```


## See also


[Chart Object](PowerPoint.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]