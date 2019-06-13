---
title: FreeformBuilder.AddNodes method (Publisher)
keywords: vbapb10.chm3276816
f1_keywords:
- vbapb10.chm3276816
ms.prod: publisher
api_name:
- Publisher.FreeformBuilder.AddNodes
ms.assetid: 29906bde-e6a6-f661-0f3f-085f39653e42
ms.date: 06/08/2019
localization_priority: Normal
---


# FreeformBuilder.AddNodes method (Publisher)

Inserts a new segment at the end of the freeform that is being created, and adds the nodes that define the segment. 

You can use this method as many times as you want to add nodes to the freeform that you are creating. When you finish adding nodes, use the **[ConvertToShape](Publisher.FreeformBuilder.ConvertToShape.md)** method to create the freeform that you just defined.


## Syntax

_expression_.**AddNodes** (_SegmentType_, _EditingType_, _X1_, _Y1_, _X2_, _Y2_, _X3_, _Y3_)

_expression_ A variable that represents a **[FreeformBuilder](Publisher.FreeformBuilder.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_SegmentType_|Required| **[MsoSegmentType](office.msosegmenttype.md)**|The type of segment to be added. Can be **msoSegmentCurve** or **msoSegmentLine**.|
|_EditingType_|Required| **[MsoEditingType](office.msoeditingtype.md)**|The editing type of the new node. Can be **msoEditingAuto** or **msoEditingCorner**.<br/><br/>If _SegmentType_ is **msoSegmentLine**, _EditingType_ must be **msoEditingAuto**; otherwise, an error occurs.|
|_X1_|Required| **Variant**|If the _EditingType_ of the new segment is **msoEditingAuto**, this argument specifies the horizontal distance from the upper-left corner of the page to the endpoint of the new segment.<br/><br/>If the _EditingType_ of the new node is **msoEditingCorner**, this argument specifies the horizontal distance from the upper-left corner of the page to the first control point for the new segment.|
|_Y1_|Required| **Variant**|If the _EditingType_ of the new segment is **msoEditingAuto**, this argument specifies the vertical distance from the upper-left corner of the page to the endpoint of the new segment.<br/><br/>If the _EditingType_ of the new node is **msoEditingCorner**, this argument specifies the vertical distance from the upper-left corner of the page to the first control point for the new segment.|
|_X2_|Optional| **Variant**|If the _EditingType_ of the new segment is **msoEditingCorner**, this argument specifies the horizontal distance from the upper-left corner of the page to the second control point for the new segment.<br/><br/>If the _EditingType_ of the new segment is **msoEditingAuto**, do not specify a value for this argument.|
|_Y2_|Optional| **Variant**|If the _EditingType_ of the new segment is **msoEditingCorner**, this argument specifies the vertical distance from the upper-left corner of the page to the second control point for the new segment.<br/><br/>If the _EditingType_ of the new segment is **msoEditingAuto**, do not specify a value for this argument.|
|_X3_|Optional| **Variant**|If the _EditingType_ of the new segment is **msoEditingCorner**, this argument specifies the horizontal distance from the upper-left corner of the page to the endpoint of the new segment.<br/><br/>If the _EditingType_ of the new segment is **msoEditingAuto**, do not specify a value for this argument.|
|_Y3_|Optional| **Variant**|If the _EditingType_ of the new segment is **msoEditingAuto**, this argument specifies the vertical distance from the upper-left corner of the page to the endpoint of the new segment.<br/><br/>If the _EditingType_ of the new segment is **msoEditingAuto**, do not specify a value for this argument.|

## Remarks

For the _X1_, _Y1_, _X2_, _Y2_, _X3_, and _Y3_ arguments, numeric values are evaluated in [points](../language/glossary/vbe-glossary.md#point); strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").

To add nodes to a freeform after it is created, use the **[Insert](Publisher.ShapeNodes.Insert.md)** method of the **ShapeNodes** collection.


## Example

This example adds a freeform with four vertices to the first page in the active publication.

```vb
' Add a new freeform object. 
With ActiveDocument.Pages(1).Shapes _ 
 .BuildFreeform(EditingType:=msoEditingCorner, _ 
 X1:=100, Y1:=100) 
 
 ' Add three more nodes and close the polygon. 
 .AddNodes SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingCorner, _ 
 X1:=200, Y1:=200, X2:=225, Y2:=250, X3:=250, Y3:=200 
 .AddNodes SegmentType:=msoSegmentCurve, _ 
 EditingType:=msoEditingAuto, X1:=200, Y1:=100 
 .AddNodes SegmentType:=msoSegmentLine, _ 
 EditingType:=msoEditingAuto, X1:=150, Y1:=50 
 .AddNodes SegmentType:=msoSegmentLine, _ 
 EditingType:=msoEditingAuto, X1:=100, Y1:=100 
 
 ' Convert the polygon to a Shape object. 
 .ConvertToShape 
End With 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]