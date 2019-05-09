---
title: Series.BarShape property (Word)
keywords: vbawd10.chm123733371
f1_keywords:
- vbawd10.chm123733371
ms.prod: word
api_name:
- Word.Series.BarShape
ms.assetid: da27d6a0-360f-dafa-3049-d9fdc2ee43ff
ms.date: 06/08/2017
localization_priority: Normal
---


# Series.BarShape property (Word)

Returns or sets the shape used for a single series in a 3D bar or column chart. Read/write  **[XlBarShape](Word.xlbarshape.md)**.


## Syntax

_expression_.**BarShape**

_expression_ A variable that represents a '[Series](Word.Series.md)' object.


## Example

The following example sets the shape used for the first series of the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).BarShape = xlConeToPoint 
 End If 
End With
```


## See also


[Series Object](Word.Series.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]