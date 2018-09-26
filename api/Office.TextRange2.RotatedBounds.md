---
title: TextRange2.RotatedBounds Method (Office)
ms.prod: office
api_name:
- Office.TextRange2.RotatedBounds
ms.assetid: e8e1b0dc-426f-c804-e91a-8cd5345186de
ms.date: 06/08/2017
---


# TextRange2.RotatedBounds Method (Office)

Gets the coordinates of the vertices of the text bounding box for the specified text range. Read-only.


## Syntax

 _expression_. `RotatedBounds`( `_X1_`, `_Y1_`, `_X2_`, `_Y2_`, `_X3_`, `_Y3_`, `_x4_`, `_y4_` )

 _expression_ An expression that returns a [TextRange2](./Office.TextRange2.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _X1_|Required|**Single**|Returns the position (in points) of the X coordinate of the first vertex of the bounding box for the text within the specified text range.|
| _Y1_|Required|**Single**|Returns the position (in points) of the Y coordinate of the first vertex of the bounding box for the text within the specified text range.|
| _X2_|Required|**Single**|Returns the position (in points) of the X coordinate of the second vertex of the bounding box for the text within the specified text range.|
| _Y2_|Required|**Single**|Returns the position (in points) of the Y coordinate of the second vertex of the bounding box for the text within the specified text range.|
| _X3_|Required|**Single**|Returns the position (in points) of the X coordinate of the third vertex of the bounding box for the text within the specified text range.|
| _Y3_|Required|**Single**|Returns the position (in points) of the Y coordinate of the third vertex of the bounding box for the text within the specified text range.|
| _x4_|Required|**Single**|Returns the position (in points) of the X coordinate of the fourth vertex of the bounding box for the text within the specified text range.|
| _y4_|Required|**Single**|Returns the position (in points) of the Y coordinate of the fourth vertex of the bounding box for the text within the specified text range.|

## Remarks

The text bounding box is not the same as the  **TextFrame2** object. The **TextFrame2** object represents the container in which the text can reside. The text bounding box represents the perimeter immediately surrounding the text.


## Example

This example uses the values returned by the arguments of the  **RotatedBounds** method to draw a freeform that has the dimensions of the text bounding box for the third word in the text range in shape one on slide one in the active presentation.


```vb
Dim x1 As Single, y1 As Single 
Dim x2 As Single, y2 As Single 
Dim x3 As Single, y3 As Single 
Dim x4 As Single, y4 As Single 
Dim myDocument As Slide 
 
Set myDocument = ActivePresentation.Slides(1) 
myDocument.Shapes(1).TextFrame2.TextRange2.Words(3).RotatedBounds _ 
 x1, y1, x2, y2, x3, y3, x4, y4 
With myDocument.Shapes.BuildFreeform(msoEditingCorner, x1, y1) 
 .AddNodes msoSegmentLine, msoEditingAuto, x2, y2 
 .AddNodes msoSegmentLine, msoEditingAuto, x3, y3 
 .AddNodes msoSegmentLine, msoEditingAuto, x4, y4 
 .AddNodes msoSegmentLine, msoEditingAuto, x1, y1 
 .ConvertToShape.ZOrder msoSendToBack 
End With 

```


## See also


[TextRange2 Object](Office.TextRange2.md)



[TextRange2 Object Members](./overview/Library-Reference/textrange2-members-office.md)

