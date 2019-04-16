---
title: Chart.Legend property (Word)
keywords: vbawd10.chm79364180
f1_keywords:
- vbawd10.chm79364180
ms.prod: word
api_name:
- Word.Chart.Legend
ms.assetid: b1ffdbfb-854c-bd65-dd63-d3b8d0547f67
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.Legend property (Word)

Returns the legend for the chart. Read-only  **[Legend](Word.Legend.md)**.


## Syntax

_expression_.**Legend**

_expression_ A variable that represents a **[Chart](Word.Chart.md)** object.


## Example

The following example enables the legend for the first chart in the active document and then sets the legend font color to blue.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart 
 .HasLegend = True 
 .Legend.Font.ColorIndex = 5 
 End With 
 End If 
End With
```


## See also


[Chart Object](Word.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]