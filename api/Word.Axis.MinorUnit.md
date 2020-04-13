---
title: Axis.MinorUnit property (Word)
keywords: vbawd10.chm113049639
f1_keywords:
- vbawd10.chm113049639
ms.prod: word
api_name:
- Word.Axis.MinorUnit
ms.assetid: 9272b2da-0067-b180-a11f-1bec0dc1a416
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.MinorUnit property (Word)

Returns or sets the minor units on the value axis. Read/write  **Double**.


## Syntax

_expression_. `MinorUnit`

_expression_ A variable that represents an **[Axis](Word.Axis.md)** object.


## Remarks

Setting this property sets the **[MinorUnitIsAuto](Word.Axis.MinorUnitIsAuto.md)** property to **False**.

Use the **[TickMarkSpacing](Word.Axis.TickLabelSpacing.md)** property to set tick-mark spacing on the category axis.


## Example

The following example sets the major and minor units for the value axis of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 With .Chart.Axes(xlValue) 
 .MajorUnit = 100 
 .MinorUnit = 20 
 End With 
 End If 
End With
```


## See also


[Axis Object](Word.Axis.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]