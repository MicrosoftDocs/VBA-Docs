---
title: Axis.ScaleType property (Word)
keywords: vbawd10.chm113049645
f1_keywords:
- vbawd10.chm113049645
ms.prod: word
api_name:
- Word.Axis.ScaleType
ms.assetid: 3b48280e-378d-81f2-133f-b5b21f63f7b1
ms.date: 06/08/2017
localization_priority: Normal
---


# Axis.ScaleType property (Word)

Returns or sets the value axis scale type. Read/write  **[XlScaleType](Word.xlscaletype.md)**.


## Syntax

_expression_. `ScaleType`

_expression_ A variable that represents an **[Axis](Word.Axis.md)** object.


## Example

The following example sets the value axis for the first chart in the active document to use a logarithmic scale.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.Axes(xlValue).ScaleType = xlScaleLogarithmic 
 End If 
End With 

```


## See also


[Axis Object](Word.Axis.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]