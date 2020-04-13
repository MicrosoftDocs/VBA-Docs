---
title: Shapes.AddCanvas method (Word)
keywords: vbawd10.chm161415193
f1_keywords:
- vbawd10.chm161415193
ms.prod: word
api_name:
- Word.Shapes.AddCanvas
ms.assetid: ff6da70f-f6ce-83f8-8e30-95b50a1f4e4f
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.AddCanvas method (Word)

Adds a drawing canvas to a document. Returns a  **[Shape](Word.Shape.md)** object that represents the drawing canvas and adds it to the **Shapes** collection.


## Syntax

_expression_. `AddCanvas`( `_Left_` , `_Top_` , `_Width_` , `_Height_` , `_Anchor_` )

_expression_ Required. A variable that represents a **[Shapes](Word.Shapes.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Left_|Required| **Single**|The position, in [points](../language/glossary/vbe-glossary.md#point), of the left edge of the drawing canvas, relative to the anchor.|
| _Top_|Required| **Single**|The position, in [points](../language/glossary/vbe-glossary.md#point), of the top edge of the drawing canvas, relative to the anchor.|
| _Width_|Required| **Single**|The width, in [points](../language/glossary/vbe-glossary.md#point), of the drawing canvas.|
| _Height_|Required| **Single**|The height, in [points](../language/glossary/vbe-glossary.md#point), of the drawing canvas.|
| _Anchor_|Optional| **Variant**|A **[Range](Word.Range.md)** object that represents the text to which the canvas is bound. If Anchor is specified, the anchor is positioned at the beginning of the first paragraph in the anchoring range. If this argument is omitted, the anchoring range is selected automatically and the canvas is positioned relative to the top and left edges of the page.|

## Return value

Shape


## Example

The following example adds a drawing canvas to a new document and formats the drawing canvas so it is inline with the text; then adds two shapes to the canvas, and formats the fill and line properties.


```vb
Sub AddInlineCanvas() 
 Dim docNew As Document 
 Dim shpCanvas As Shape 
 
 Set docNew = Documents.Add 
 
 'Add a drawing canvas to the new document 
 Set shpCanvas = docNew.Shapes.AddCanvas( _ 
 Left:=150, Top:=150, Width:=70, Height:=70) 
 shpCanvas.WrapFormat.Type = wdWrapInline 
 
 'Add shapes to drawing canvas 
 With shpCanvas.CanvasItems 
 .AddShape msoShapeHeart, Left:=10, _ 
 Top:=10, Width:=50, Height:=60 
 .AddLine BeginX:=0, BeginY:=0, _ 
 EndX:=70, EndY:=70 
 End With 
 With shpCanvas 
 .CanvasItems(1).Fill.ForeColor _ 
 .RGB = RGB(Red:=255, Green:=0, Blue:=0) 
 .CanvasItems(2).Line _ 
 .EndArrowheadStyle = msoArrowheadTriangle 
 End With 
End Sub
```


## See also


[Shapes Collection Object](Word.shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]