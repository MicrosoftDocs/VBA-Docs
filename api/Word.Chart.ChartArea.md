---
title: Chart.ChartArea property (Word)
keywords: vbawd10.chm79364157
f1_keywords:
- vbawd10.chm79364157
ms.prod: word
api_name:
- Word.Chart.ChartArea
ms.assetid: b16d78c0-7663-3ef9-c17a-02e7a024b344
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.ChartArea property (Word)

Returns the complete chart area for the chart. Read-only  **[ChartArea](Word.ChartArea.md)**.


## Syntax

_expression_. `ChartArea`

_expression_ A variable that represents a **[Chart](Word.Chart.md)** object.


## Example

The following example sets the chart area interior color of the first chart in the active document to red and sets the border color to blue.


```vb
With ActiveDocument.InlineShapes(1).Chart.ChartArea 
 .Interior.ColorIndex = 3 
 .Border.ColorIndex = 5 
End With
```


## See also


[Chart Object](Word.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]