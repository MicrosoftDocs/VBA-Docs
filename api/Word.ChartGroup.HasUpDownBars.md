---
title: ChartGroup.HasUpDownBars property (Word)
keywords: vbawd10.chm263454738
f1_keywords:
- vbawd10.chm263454738
ms.prod: word
api_name:
- Word.ChartGroup.HasUpDownBars
ms.assetid: 9c39f015-f8cc-633c-54a0-b68fc420d8f6
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartGroup.HasUpDownBars property (Word)

 **True** if a line chart has up and down bars. Read/write **Boolean**.


## Syntax

_expression_.**HasUpDownBars**

_expression_ A variable that represents a **[ChartGroup](Word.ChartGroup.md)** object.


## Remarks

This property applies only to line charts. 


## Example

The following example enables up and down bars for chart group one of the first chart in the active document and then sets their colors. You should run the example on a 2D line chart that contains two series that cross each other at one or more data points.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.ChartGroups(1) 
 .HasUpDownBars = True 
 .DownBars.Interior.ColorIndex = 3 
 .UpBars.Interior.ColorIndex = 5 
 End With 
 End If 
End With
```


## See also


[ChartGroup Object](Word.ChartGroup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]