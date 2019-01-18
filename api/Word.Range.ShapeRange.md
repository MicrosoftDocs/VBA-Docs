---
title: Range.ShapeRange property (Word)
keywords: vbawd10.chm157155639
f1_keywords:
- vbawd10.chm157155639
ms.prod: word
api_name:
- Word.Range.ShapeRange
ms.assetid: b8e6e1f7-d29a-5fb5-8d00-22b3907d6f54
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.ShapeRange property (Word)

Returns a  **[ShapeRange](Word.shapeRange.md)** collection that represents all the **Shape** objects in the specified range. Read-only.


## Syntax

 _expression_. `ShapeRange`

 _expression_ A variable that represents a '[Range](Word.Range.md)' object.


## Remarks

A shape range can contain drawings, shapes, pictures, OLE objects, ActiveX controls, text objects, and callouts.


## Example

The following example sets the fill foreground color to purple for all the shapes in the active document.


```vb
ActiveDocument.Content.ShapeRange.Fill.ForeColor.RGB = _ 
 RGB(255, 0, 255)
```


## See also


[Range Object](Word.Range.md)

