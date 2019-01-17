---
title: ChartArea.ClearContents method (Word)
keywords: vbawd10.chm160039025
f1_keywords:
- vbawd10.chm160039025
ms.prod: word
api_name:
- Word.ChartArea.ClearContents
ms.assetid: d6642767-e8f5-8834-ec8f-e78ae2994a7b
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartArea.ClearContents method (Word)

Clears the data from a chart but leaves the formatting.


## Syntax

 _expression_. `ClearContents`

 _expression_ A variable that represents a '[ChartArea](Word.ChartArea.md)' object.


## Example

The following example clears the chart data from the first chart in the active document but leaves the formatting intact.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.ChartArea.ClearContents 
 End If 
End With
```


## See also


[ChartArea Object](Word.ChartArea.md)

