---
title: Axis.TickMarkSpacing property (Word)
keywords: vbawd10.chm113049653
f1_keywords:
- vbawd10.chm113049653
ms.prod: word
api_name:
- Word.Axis.TickMarkSpacing
ms.assetid: 926ae9ad-0c5a-c61a-55fb-1503a2edf593
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.TickMarkSpacing property (Word)

Returns or sets the number of categories or series between tick marks. Read/write  **Long**.


## Syntax

_expression_. `TickMarkSpacing`

_expression_ A variable that represents an **[Axis](Word.Axis.md)** object.


## Remarks

This property applies only to category and series axes. It can be a value from 1 through 31999. 

Use the **[MajorUnit](Word.Axis.MajorUnit.md)** and **[MinorUnit](Word.Axis.MinorUnit.md)** properties to set tick-mark spacing on the value axis.


## Example

The following example sets the number of categories between tick marks on the category axis of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Axes(xlCategory).TickMarkSpacing = 10 
 End If 
End With
```


## See also


[Axis Object](Word.Axis.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]