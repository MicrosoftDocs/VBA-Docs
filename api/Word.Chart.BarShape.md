---
title: Chart.BarShape property (Word)
keywords: vbawd10.chm79364168
f1_keywords:
- vbawd10.chm79364168
ms.prod: word
api_name:
- Word.Chart.BarShape
ms.assetid: e29af332-162c-4a9e-0281-f546bd00f27c
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.BarShape property (Word)

Returns or sets the shape used for every series in a 3D bar or column chart. Read/write  **[XlBarShape](Word.xlbarshape.md)**.


## Syntax

_expression_.**BarShape**

_expression_ A variable that represents a **[Chart](Word.Chart.md)** object.


## Example

The following example sets the shape used with the first series of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.BarShape = xlConeToPoint 
 End If 
End With
```


## See also


[Chart Object](Word.Chart.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]