---
title: ShapeRange.AutoShapeType property (Excel)
keywords: vbaxl10.chm640098
f1_keywords:
- vbaxl10.chm640098
ms.prod: excel
api_name:
- Excel.ShapeRange.AutoShapeType
ms.assetid: de4c8273-2804-012c-38a0-7689aa54b02e
ms.date: 06/08/2017
localization_priority: Normal
---


# ShapeRange.AutoShapeType property (Excel)

Returns or sets the shape type for the specified  **[Shape](Excel.Shape.md)** or **[ShapeRange](Excel.ShapeRange.md)** object, which must represent an AutoShape other than a line, freeform drawing, or connector. Read/write **[MsoAutoShapeType](Office.MsoAutoShapeType.md)**.


## Syntax

_expression_. `AutoShapeType`

_expression_ A variable that represents a [ShapeRange](./Excel.ShapeRange.md) object.


## Remarks

When you change the type of a shape, the shape retains its size, color, and other attributes.

Use the  **[Type](Excel.ConnectorFormat.Type.md)** property of the **[ConnectorFormat](Excel.ConnectorFormat.md)** object to set or return the connector type.


## Example

This example replaces all 16-point stars with 32-point stars in  `myDocument`.


```vb
Set myDocument = Worksheets(1) 
For Each s In myDocument.Shapes 
    If s.AutoShapeType = msoShape16pointStar Then 
        s.AutoShapeType = msoShape32pointStar 
    End If 
Next
```


## See also


[ShapeRange Object](Excel.ShapeRange.md)

