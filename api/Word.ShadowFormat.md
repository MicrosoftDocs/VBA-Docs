---
title: ShadowFormat object (Word)
keywords: vbawd10.chm2508
f1_keywords:
- vbawd10.chm2508
ms.prod: word
api_name:
- Word.ShadowFormat
ms.assetid: 2a179f0b-ec18-c3dd-dd73-51b18f42e0e2
ms.date: 06/08/2017
localization_priority: Normal
---


# ShadowFormat object (Word)

Represents shadow formatting for a shape.


## Remarks

Use the **Shadow** property to return a **ShadowFormat** object. The following example adds a shadowed rectangle to the active document. The semitransparent, blue shadow is offset 5 points to the right of the rectangle and 3 points above it.


```vb
With ActiveDocument.Shapes _ 
 .AddShape(msoShapeRectangle, 50, 50, 100, 200).Shadow 
 .ForeColor.RGB = RGB(0, 0, 128) 
 .OffsetX = 5 
 .OffsetY = -3 
 .Transparency = 0.5 
 .Visible = True 
End With
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]