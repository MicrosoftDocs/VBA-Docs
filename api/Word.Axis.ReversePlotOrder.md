---
title: Axis.ReversePlotOrder property (Word)
keywords: vbawd10.chm113049643
f1_keywords:
- vbawd10.chm113049643
ms.prod: word
api_name:
- Word.Axis.ReversePlotOrder
ms.assetid: 663a1268-d7ed-0af4-afa6-1637a94f4525
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.ReversePlotOrder property (Word)

 **True** if Microsoft Word plots data points from last to first. Read/write **Boolean**.


## Syntax

_expression_.**ReversePlotOrder**

_expression_ A variable that represents an **[Axis](Word.Axis.md)** object.


## Remarks

You cannot use this property on radar charts.


## Example

The following example plots data points from last to first on the value axis for the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Axes(xlValue).ReversePlotOrder = True 
 End If 
End With
```


## See also


[Axis Object](Word.Axis.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]