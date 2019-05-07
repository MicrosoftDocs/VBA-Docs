---
title: ShapeRange.HorizontalFlip property (PowerPoint)
keywords: vbapp10.chm548025
f1_keywords:
- vbapp10.chm548025
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.HorizontalFlip
ms.assetid: 4c41e250-2a8f-3eab-3244-0910fb43362e
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.HorizontalFlip property (PowerPoint)

Returns whether the specified shape is flipped around the horizontal axis. Read-only.


## Syntax

_expression_. `HorizontalFlip`

_expression_ A variable that represents a **[ShapeRange](PowerPoint.ShapeRange.md)** object.


## Return value

MsoTriState


## Remarks

The value of the  **HorizontalFlip** property can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**| The specified shape is not flipped around the horizontal axis.|
|**msoTrue**| The specified shape is flipped around the horizontal axis.|

## Example

This example restores each shape on _myDocument_ to its original state, if it is been flipped horizontally or vertically.


```vb
Set myDocument = ActivePresentation.Slides(1)

For Each s In myDocument.Shapes

    If s.HorizontalFlip Then s.Flip msoFlipHorizontal

    If s.VerticalFlip Then s.Flip msoFlipVertical

Next
```


## See also


[ShapeRange Object](PowerPoint.ShapeRange.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]