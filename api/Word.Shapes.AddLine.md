---
title: Shapes.AddLine method (Word)
keywords: vbawd10.chm161415182
f1_keywords:
- vbawd10.chm161415182
ms.prod: word
api_name:
- Word.Shapes.AddLine
ms.assetid: d1c609c3-d5d1-80e8-4f95-184a9a536feb
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.AddLine method (Word)

Adds a line to a drawing canvas.


## Syntax

_expression_.**AddLine** (_BeginX_, _BeginY_, _EndX_, _EndY_)

_expression_ Required. A variable that represents a **[Shapes](Word.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _BeginX_|Required| **Single**|The horizontal position, measured in points, of the line's starting point, relative to the drawing canvas.|
| _BeginY_|Required| **Single**|The vertical position, measured in points, of the line's starting point, relative to the drawing canvas.|
| _EndX_|Required| **Single**|The horizontal position, measured in points, of the line's endpoint, relative to the drawing canvas.|
| _EndY_|Required| **Single**|The vertical position, measured in points, of the line's endpoint, relative to the drawing canvas.|

## Return value

 **[Shape](Word.Shape.md)**


## Remarks

To create an arrow, use the **Line** property to format a line.


## Example

This example adds a purple line with an arrow to a new drawing canvas.


```vb
Sub NewCanvasLine() 
 Dim shpCanvas As Shape 
 Dim shpLine As Shape 
 
 'Add new drawing canvas to the active document 
 Set shpCanvas = ActiveDocument.Shapes _ 
 .AddCanvas(Left:=100, Top:=75, _ 
 Width:=150, Height:=200) 
 
 'Add a line to the drawing canvas 
 Set shpLine = shpCanvas.CanvasItems.AddLine( _ 
 BeginX:=25, BeginY:=25, EndX:=150, EndY:=150) 
 
 'Add an arrow to the line and sets the color to purple 
 With shpLine.Line 
 .BeginArrowheadStyle = msoArrowheadDiamond 
 .BeginArrowheadWidth = msoArrowheadWide 
 .ForeColor.RGB = RGB(Red:=150, Green:=0, Blue:=255) 
 End With 
End Sub
```


## See also


[Shapes Collection Object](Word.shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]