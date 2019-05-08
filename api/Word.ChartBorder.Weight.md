---
title: ChartBorder.Weight property (Word)
keywords: vbawd10.chm61014022
f1_keywords:
- vbawd10.chm61014022
ms.prod: word
api_name:
- Word.ChartBorder.Weight
ms.assetid: f1fc8001-0437-0e4c-d158-8aed3d254360
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartBorder.Weight property (Word)

Returns or sets the weight of the border. Read/write  **[XlBorderWeight](Word.xlborderweight.md)**.


## Syntax

_expression_.**Weight**

_expression_ A variable that represents a '[ChartBorder](Word.ChartBorder.md)' object.


## Example

The following example sets the border weight for the value axis of the first chart in the active document to medium.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Axes(xlValue).Border.Weight = xlMedium 
 End If 
End With
```


## See also


[ChartBorder Object](Word.ChartBorder.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]