---
title: CanvasShapes.AddCallout method (Word)
keywords: vbawd10.chm7536650
f1_keywords:
- vbawd10.chm7536650
ms.prod: word
api_name:
- Word.CanvasShapes.AddCallout
ms.assetid: 87affac3-523e-165f-690a-75bd9e0b9961
ms.date: 06/08/2017
localization_priority: Normal
---


# CanvasShapes.AddCallout method (Word)

Adds a borderless line callout to a drawing canvas. Returns a  **[Shape](Word.Shape.md)** object that represents the callout.


## Syntax

_expression_.**AddCallout** (_Type_, _Left_, _Top_, _Width_, _Height_)

_expression_ Required. A variable that represents a **[CanvasShapes](Word.CanvasShapes.md)** collection.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **MsoCallout**|The type of callout.|
| _Left_|Required| **Single**|The position, in [points](../language/glossary/vbe-glossary.md#point), of the left edge of the callout's bounding box.|
| _Top_|Required| **Single**|The position, in [points](../language/glossary/vbe-glossary.md#point), of the top edge of the callout's bounding box.|
| _Width_|Required| **Single**|The width, in [points](../language/glossary/vbe-glossary.md#point), of the callout's bounding box.|
| _Height_|Required| **Single**|The height, in [points](../language/glossary/vbe-glossary.md#point), of the callout's bounding box.|

## Remarks

You can insert a greater variety of callouts, such as balloons and clouds, using the **AddShape** method.


## Example

This example adds a callout to a newly created drawing canvas.


```vb
Sub NewCanvasCallout() 
 Dim shpCanvas As Shape 
 
 'Add drawing canvas to the active document 
 Set shpCanvas = ActiveDocument.Shapes.AddCanvas _ 
 (Left:=150, Top:=150, Width:=200, Height:=300) 
 
 'Add callout to the drawing canvas 
 shpCanvas.CanvasItems.AddCallout _ 
 Type:=msoCalloutTwo, Left:=100, _ 
 Top:=40, Width:=150, Height:=75 
End Sub
```


## See also


[CanvasShapes Collection](Word.CanvasShapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]