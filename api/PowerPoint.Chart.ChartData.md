---
title: Chart.ChartData property (PowerPoint)
keywords: vbapp10.chm684011
f1_keywords:
- vbapp10.chm684011
ms.prod: powerpoint
api_name:
- PowerPoint.Chart.ChartData
ms.assetid: 16262f71-13cd-a023-35df-2ca6bd017e3b
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.ChartData property (PowerPoint)

Returns information about the linked or embedded data associated with a chart. Read-only  **[ChartData](PowerPoint.ChartData.md)**.


## Syntax

_expression_. `ChartData`

_expression_ A variable that represents a **[Chart](PowerPoint.Chart.md)** object.


## Example




> [!NOTE] 
> Although the following code applies to Microsoft Word, you can readily modify it to apply to PowerPoint.

The following example uses the  **[Activate](PowerPoint.ChartData.Activate.md)** method to display the data associated with the first chart in the active document.




```vb
With ActiveDocument.InlineShapes(1).Chart.ChartData

    .Activate

End With
```


## See also


[Chart Object](PowerPoint.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]