---
title: Chart.ChartData property (Word)
keywords: vbawd10.chm79364189
f1_keywords:
- vbawd10.chm79364189
ms.prod: word
api_name:
- Word.Chart.ChartData
ms.assetid: d8234dd3-148f-b69a-8a4e-f22474080eab
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.ChartData property (Word)

Returns information about the linked or embedded data associated with a chart. Read-only  **[ChartData](Word.ChartData.md)**.


## Syntax

_expression_. `ChartData`

_expression_ A variable that represents a **[Chart](Word.Chart.md)** object.


## Example

The following example uses the  **[Activate](Word.ChartData.Activate.md)** method to display the data associated with the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1).Chart.ChartData 
 .Activate 
End With
```


## See also


[Chart Object](Word.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]