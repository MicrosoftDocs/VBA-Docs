---
title: ShapeNodes.Insert method (Publisher)
keywords: vbapb10.chm3473426
f1_keywords:
- vbapb10.chm3473426
ms.prod: publisher
api_name:
- Publisher.ShapeNodes.Insert
ms.assetid: c78ceefe-db9f-4af0-2e76-2ab1e4dc74b8
ms.date: 06/14/2019
localization_priority: Normal
---


# ShapeNodes.Insert method (Publisher)

Inserts a new segment after the specified node of the freeform drawing.


## Syntax

_expression_.**Insert** (_Index_, _SegmentType_, _EditingType_, _X1_, _Y1_, _X2_, _Y2_, _X3_, _Y3_)

_expression_ A variable that represents a **[ShapeNodes](Publisher.ShapeNodes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_|Required| **Long**|The number of the node after which the new node is to be inserted.|
|_SegmentType_|Required| **[MsoSegmentType](office.msosegmenttype.md)**|The type of segment to be added. Can be one of the **MsoSegmentType** constants.|
|_EditingType_|Required| **[MsoEditingType](Office.MsoEditingType.md)**|The editing type of the new node. Can be one of the **MsoEditingType** constants.|
|_X1_|Required| **Variant**|If the _EditingType_ of the new segment is **msoEditingAuto**, this argument specifies the horizontal distance from the upper-left corner of the page to the endpoint of the new segment.<br/><br/>If the _EditingType_ of the new node is **msoEditingCorner**, this argument specifies the horizontal distance from the upper-left corner of the page to the first control point for the new segment.|
|_Y1_|Required| **Variant**|If the _EditingType_ of the new segment is **msoEditingAuto**, this argument specifies the vertical distance from the upper-left corner of the page to the endpoint of the new segment.<br/><br/>If the _EditingType_ of the new node is **msoEditingCorner**, this argument specifies the vertical distance from the upper-left corner of the page to the first control point for the new segment.|
|_X2_|Optional| **Variant**|If the _EditingType_ of the new segment is **msoEditingCorner**, this argument specifies the horizontal distance from the upper-left corner of the page to the second control point for the new segment.<br/><br/>If the _EditingType_ of the new segment is **msoEditingAuto**, do not specify a value for this argument.|
|_Y2_|Optional| **Variant**|If the _EditingType_ of the new segment is **msoEditingCorner**, this argument specifies the vertical distance from the upper-left corner of the page to the second control point for the new segment.<br/><br/>If the _EditingType_ of the new segment is **msoEditingAuto**, do not specify a value for this argument.|
|_X3_|Optional| **Variant**|If the _EditingType_ of the new segment is **msoEditingCorner**, this argument specifies the horizontal distance from the upper-left corner of the page to the endpoint of the new segment.<br/><br/>If the _EditingType_ of the new segment is **msoEditingAuto**, do not specify a value for this argument.|
|_Y3_|Optional| **Variant**|If the _EditingType_ of the new segment is **msoEditingCorner**, this argument specifies the vertical distance from the upper-left corner of the page to the endpoint of the new segment.<br/><br/>If the _EditingType_ of the new segment is **msoEditingAuto**, do not specify a value for this argument.|

## Remarks

For the _X1_, _Y1_, _X2_, _Y2_, _X3_, and _Y3_ arguments, numeric values are evaluated in [points](../language/glossary/vbe-glossary.md#point); strings can be in any units supported by Publisher (for example, "2.5 in"). 


## Example

This example adds a smooth node with a curved segment after node four in the third shape in the active publication. The shape must be a freeform drawing with at least four nodes.

```vb
With ActiveDocument.Pages(1).Shapes(3).Nodes 
 .Insert Index:=4, _ 
 SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingAuto, _ 
 X1:=210, Y1:=100 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]