---
title: ShapeNodes.Insert method (Excel)
keywords: vbaxl10.chm112008
f1_keywords:
- vbaxl10.chm112008
ms.prod: excel
api_name:
- Excel.ShapeNodes.Insert
ms.assetid: b4f7e695-2102-5cbd-2d6b-bc167407cc0f
ms.date: 05/14/2019
localization_priority: Normal
---


# ShapeNodes.Insert method (Excel)

Inserts a node into a freeform shape.


## Syntax

_expression_.**Insert** (_Index_, _SegmentType_, _EditingType_, _X1_, _Y1_, _X2_, _Y2_, _X3_, _Y3_)

_expression_ A variable that represents a **[ShapeNodes](Excel.ShapeNodes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Integer**| **Long**. The number of the shape node after which to insert a new node.|
| _SegmentType_|Required| **[MsoSegmentType](Office.MsoSegmentType.md)**|The segment type.|
| _EditingType_|Required| **[MsoEditingType](Office.MsoEditingType.md)**|The editing type.|
| _X1_|Required| **Single**|If the _EditingType_ of the new segment is **msoEditingAuto**, this argument specifies the horizontal distance, measured in [points](../language/glossary/vbe-glossary.md#point), from the upper-left corner of the document to the end point of the new segment.<br/><br/>If the _EditingType_ of the new node is **msoEditingCorner**, this argument specifies the horizontal distance, measured in points, from the upper-left corner of the document to the first control point for the new segment.|
| _Y1_|Required| **Single**|If the _EditingType_ of the new segment is **msoEditingAuto**, this argument specifies the vertical distance, measured in points, from the upper-left corner of the document to the end point of the new segment.<br/><br/>If the _EditingType_ of the new node is **msoEditingCorner**, this argument specifies the vertical distance, measured in points, from the upper-left corner of the document to the first control point for the new segment.|
| _X2_|Required| **Single**|If the _EditingType_ of the new segment is **msoEditingCorner**, this argument specifies the horizontal distance, measured in points, from the upper-left corner of the document to the second control point for the new segment.<br/><br/>If the _EditingType_ of the new segment is **msoEditingAuto**, don't specify a value for this argument.|
| _Y2_|Required| **Single**|If the _EditingType_ of the new segment is **msoEditingCorner**, this argument specifies the vertical distance, measured in points, from the upper-left corner of the document to the second control point for the new segment.<br/><br/>If the _EditingType_ of the new segment is **msoEditingAuto**, don't specify a value for this argument.|
| _X3_|Required| **Single**|If the _EditingType_ of the new segment is **msoEditingCorner**, this argument specifies the horizontal distance, measured in points, from the upper-left corner of the document to the end point of the new segment.<br/><br/>If the _EditingType_ of the new segment is **msoEditingAuto**, don't specify a value for this argument.|
| _Y3_|Required| **Single**|If the _EditingType_ of the new segment is **msoEditingCorner**, this argument specifies the vertical distance, measured in points, from the upper-left corner of the document to the end point of the new segment.<br/><br/>If the _EditingType_ of the new segment is **msoEditingAuto**, don't specify a value for this argument.|

## Example

This example selects the third shape in the active document, checks whether the shape is a Freeform object, and if it is, inserts a node. This example assumes three shapes exist on the active worksheet.

```vb
Sub InsertShapeNode() 
    ActiveSheet.Shapes(3).Select 
    With Selection.ShapeRange 
        If .Type = msoFreeform Then 
            .Nodes.Insert _ 
                Index:=3, SegmentType:=msoSegmentCurve, _ 
                EditingType:=msoEditingSymmetric, X1:=35, Y1:=100 
            .Fill.ForeColor.RGB = RGB(0, 0, 200) 
            .Fill.Visible = msoTrue 
        Else 
            MsgBox "This shape is not a Freeform object." 
        End If 
    End With 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]