---
title: Axis.MaximumScale property (Word)
keywords: vbawd10.chm113049628
f1_keywords:
- vbawd10.chm113049628
ms.prod: word
api_name:
- Word.Axis.MaximumScale
ms.assetid: cfd12a67-ef8b-d92c-a9c1-74353754498e
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.MaximumScale property (Word)

Returns or sets the maximum value on the value axis. Read/write  **Double**.


## Syntax

_expression_. `MaximumScale`

_expression_ A variable that represents an **[Axis](Word.Axis.md)** object.


## Remarks

Setting this property sets the **[MaximumScaleIsAuto](Word.Axis.MaximumScaleIsAuto.md)** property to **False**.


## Example

The following example sets the minimum and maximum values for the value axis of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlValue) 
 .MinimumScale = 10 
 .MaximumScale = 120 
 End With 
 End If 
End With
```


## See also


[Axis Object](Word.Axis.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]