---
title: ShapeRange.VerticalFlip property (PowerPoint)
keywords: vbapp10.chm548039
f1_keywords:
- vbapp10.chm548039
ms.prod: powerpoint
api_name:
- PowerPoint.ShapeRange.VerticalFlip
ms.assetid: 868657a8-72c6-896d-6a6f-f9adbbc88a59
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.VerticalFlip property (PowerPoint)

Determines whether the specified shape is flipped around the vertical axis. Read-only.


## Syntax

_expression_.**VerticalFlip**

_expression_ A variable that represents a **[ShapeRange](PowerPoint.ShapeRange.md)** object.


## Return value

MsoTriState


## Remarks

The value of the  **VerticalFlip** property can be one of these **MsoTriState** constants.



|Constant|Description|
|:-----|:-----|
|**msoFalse**|The specified shape is not flipped around the vertical axis.|
|**msoTrue**| The specified shape is flipped around the vertical axis.|

## Example

This example restores each shape on _myDocument_ to its original state if it is been flipped horizontally or vertically.


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