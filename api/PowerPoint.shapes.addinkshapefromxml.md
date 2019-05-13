---
title: Shapes.AddInkShapeFromXML method (PowerPoint)
ms.assetid: 88a395ac-b11e-d42e-f4b4-b41bf1d1347e
ms.date: 06/08/2017
ms.prod: powerpoint
localization_priority: Normal
---


# Shapes.AddInkShapeFromXML method (PowerPoint)

Creates an ink shape. Returns a [Shape](PowerPoint.Shape.md) object that represents the new ink shape.


## Syntax

_expression_. `AddInkShapeFromXML`( _InkXML_,  _InkXML_,  _Left_,  _Top_,  _Width_,  _Height_)

_expression_ A variable that represents a **[Shapes](PowerPoint.Shapes.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _linkXML_|Required|**String**|The string that contains the InkActionML of the ink to create.|
| _Left_|Required|**Single**|The position, measured in points, of the left edge of the ink shape relative to the left edge of the slide.|
| _Top_|Required|**Single**|The position, measured in points, of the top edge of the ink shape relative to the top edge of the slide.|
| _Width_|Optional|**Single**| The width of the ink shape, measured in points. If this parameter is not specified, the width is calculated based off of the InkActionML.|
| _Height_|Optional|**Single**|The height of the ink shape, measured in points. If this parameter is not specified, the hight is calculated based off of the InkActionML.|

## Return value

A [Shape](PowerPoint.Shape.md) object that represents the newly-added ink shape.


## See also


[Shape](PowerPoint.Shape.md)
[Shapes Object](PowerPoint.Shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]