---
title: Shapes.AddShape method (PowerPoint)
keywords: vbapp10.chm543012
f1_keywords:
- vbapp10.chm543012
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.AddShape
ms.assetid: 2bc6cce5-3461-61ff-083d-bd36ee71cb59
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.AddShape method (PowerPoint)

Creates an AutoShape. Returns a **[Shape](PowerPoint.Shape.md)** object that represents the new AutoShape.


## Syntax

_expression_. `AddShape`( `_Type_`, `_Left_`, `_Top_`, `_Width_`, `_Height_` )

_expression_ A variable that represents a **[Shapes](PowerPoint.Shapes.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required|**[MsoAutoShapeType](Office.MsoAutoShapeType.md)**|Specifies the type of AutoShape to create.|
| _Left_|Required|**Single**|The position, measured in points, of the left edge of the AutoShape relative to the left edge of the slide.|
| _Top_|Required|**Single**|The position, measured in points, of the top edge of the AutoShape relative to the top edge of the slide.|
| _Width_|Required|**Single**|The width of the AutoShape, measured in points.|
| _Height_|Required|**Single**|The height of the AutoShape, measured in points.|

## Return value

Shape


## Remarks

To change the type of an AutoShape that you've added, set the  **AutoShapeType** property.


## Example

This example adds a rectangle to _myDocument_.


```vb
Set myDocument = ActivePresentation.Slides(1) 
myDocument.Shapes.AddShape Type:=msoShapeRectangle, _ 
    Left:=50, Top:=50, Width:=100, Height:=200
```


## See also


[Shapes Object](PowerPoint.Shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
