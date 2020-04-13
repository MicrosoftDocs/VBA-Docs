---
title: HorizontalLineFormat object (Word)
keywords: vbawd10.chm2526
f1_keywords:
- vbawd10.chm2526
ms.prod: word
api_name:
- Word.HorizontalLineFormat
ms.assetid: 55296fc7-9b7e-dcdb-00e0-901015cf0efb
ms.date: 06/08/2017
localization_priority: Normal
---


# HorizontalLineFormat object (Word)

Represents horizontal line formatting.


## Remarks

Use the **HorizontalLineFormat** property to return a **HorizontalLineFormat** object. This example sets the alignment for a new horizontal line.


```vb
Selection.InlineShapes.AddHorizontalLineStandard 
ActiveDocument.InlineShapes(1) _ 
 .HorizontalLineFormat.Alignment = _ 
 wdHorizontalLineAlignLeft
```

This example adds a horizontal line without any 3D shading.




```vb
Selection.InlineShapes.AddHorizontalLineStandard 
ActiveDocument.InlineShapes(1) _ 
 .HorizontalLineFormat.NoShade = True
```

This example adds a horizontal line and sets its length to 50% of the window width.




```vb
Selection.InlineShapes.AddHorizontalLineStandard 
ActiveDocument.InlineShapes(1) _ 
 .HorizontalLineFormat.PercentWidth = 50
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]