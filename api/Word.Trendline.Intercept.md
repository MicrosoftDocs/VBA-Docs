---
title: Trendline.Intercept property (Word)
keywords: vbawd10.chm26345658
f1_keywords:
- vbawd10.chm26345658
ms.prod: word
api_name:
- Word.Trendline.Intercept
ms.assetid: d1b3c93b-4af4-96cf-c6ed-27a04d7204c2
ms.date: 06/08/2017
localization_priority: Normal
---


# Trendline.Intercept property (Word)

Returns or sets the point where the trendline crosses the value axis. Read/write  **Double**.


## Syntax

_expression_. `Intercept`

_expression_ A variable that represents a '[Trendline](Word.Trendline.md)' object.


## Remarks

Setting this property sets the **[InterceptIsAuto](Word.Trendline.InterceptIsAuto.md)** property to **False**.


## Example

The following example sets trendline one for the first chart in the active document to cross the value axis at 5. You should run the example on a 2D column chart that contains a single series that has a trendline.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.SeriesCollection(1).Trendlines(1).Intercept = 5 
 End If 
End With
```


## See also


[Trendline Object](Word.Trendline.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]