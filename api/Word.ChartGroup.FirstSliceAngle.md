---
title: ChartGroup.FirstSliceAngle property (Word)
keywords: vbawd10.chm263454726
f1_keywords:
- vbawd10.chm263454726
ms.prod: word
api_name:
- Word.ChartGroup.FirstSliceAngle
ms.assetid: 0b5b9e0b-1210-6fc6-9e2c-2913cdb552cc
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartGroup.FirstSliceAngle property (Word)

Returns or sets the angle, in degrees (clockwise from vertical), of the first pie-chart or doughnut-chart slice. Read/write  **Long**.


## Syntax

_expression_.**FirstSliceAngle**

_expression_ A variable that represents a **[ChartGroup](Word.ChartGroup.md)** object.


## Remarks

This property applies only to pie, 3D pie, and doughnut charts. It can be a value from 0 through 360. 


## Example

The following example sets the angle for the first slice in chart group one for the first chart in the active document. You should run the example on a 2D pie chart.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.ChartGroups(1).FirstSliceAngle = 15 
 End If 
End With 

```


## See also


[ChartGroup Object](Word.ChartGroup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]