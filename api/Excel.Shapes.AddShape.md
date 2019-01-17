---
title: Shapes.AddShape method (Excel)
keywords: vbaxl10.chm638084
f1_keywords:
- vbaxl10.chm638084
ms.prod: excel
api_name:
- Excel.Shapes.AddShape
ms.assetid: 5d08e6d5-2875-795a-8fe1-f4032d4d3fc0
ms.date: 06/08/2017
localization_priority: Priority
---


# Shapes.AddShape method (Excel)

Returns a  **[Shape](Excel.Shape.md)** object that represents the new AutoShape in a worksheet.


## Syntax

_expression_. `AddShape`( `_Type_` , `_Left_` , `_Top_` , `_Width_` , `_Height_` )

_expression_ A variable that represents a [Shapes](./Excel.Shapes.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **[MsoAutoShapeType](Office.MsoAutoShapeType.md)**|Specifies the type of AutoShape to create.|
| _Left_|Required| **Single**|The position (in points) of the upper-left corner of the AutoShape's bounding box relative to the upper-left corner of the document.|
| _Top_|Required| **Single**|The position (in points) of the upper-left corner of the AutoShape's bounding box relative to the upper-left corner of the document.|
| _Width_|Required| **Single**|The width of the AutoShape's bounding box, in points.|
| _Height_|Required| **Single**|The height of the AutoShape's bounding box, in points.|

## Return value

Shape


## Remarks

To change the type of an AutoShape that you?ve added, set the  **[AutoShapeType](Excel.Shape.AutoShapeType.md)** property.


## Example

This example adds a rectangle to  `myDocument`.


```vb
Set myDocument = Worksheets(1) 
myDocument.Shapes.AddShape msoShapeRectangle, 50, 50, 100, 200
```


## See also


[Shapes Object](Excel.Shapes.md)

