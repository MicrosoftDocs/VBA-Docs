---
title: Shape.AutoShapeType property (Excel)
keywords: vbaxl10.chm636091
f1_keywords:
- vbaxl10.chm636091
ms.prod: excel
api_name:
- Excel.Shape.AutoShapeType
ms.assetid: 3fdc36be-dd08-4fa1-8cef-a5ecf913eae0
ms.date: 06/08/2017
localization_priority: Normal
---


# Shape.AutoShapeType property (Excel)

Returns or sets the shape type for the specified  **[Shape](Excel.Shape.md)** or **[ShapeRange](Excel.ShapeRange.md)** object, which must represent an AutoShape other than a line, freeform drawing, or connector. Read/write **[MsoAutoShapeType](Office.MsoAutoShapeType.md)**.


## Syntax

_expression_. `AutoShapeType`

_expression_ A variable that represents a [Shape](Excel.Shape.md) object.


## Remarks

When you change the type of a shape, the shape retains its size, color, and other attributes.

Use the  **[Type](Excel.ConnectorFormat.Type.md)** property of the **[ConnectorFormat](Excel.ConnectorFormat.md)** object to set or return the connector type.


## Example

This example replaces all 16-point stars with 32-point stars in  _myDocument_.


```vb
Set myDocument = Worksheets(1) 
For Each s In myDocument.Shapes 
    If s.AutoShapeType = msoShape16pointStar Then 
        s.AutoShapeType = msoShape32pointStar 
    End If 
Next
```


## See also


[Shape Object](Excel.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
