---
title: FreeformBuilder.AddNodes Method (PowerPoint)
keywords: vbapp10.chm546002
f1_keywords:
- vbapp10.chm546002
ms.prod: powerpoint
api_name:
- PowerPoint.FreeformBuilder.AddNodes
ms.assetid: 4022d4cd-796b-8917-7265-d97bff5282ef
ms.date: 06/08/2017
localization_priority: Normal
---


# FreeformBuilder.AddNodes Method (PowerPoint)

Inserts a new segment at the end of the freeform that's being created, and adds the nodes that define the segment. You can use this method as many times as you want to add nodes to the freeform you are creating. When you finish adding nodes, use the  **[ConvertToShape](PowerPoint.FreeformBuilder.ConvertToShape.md)** method to create the freeform you've just defined. To add nodes to a freeform after it is been created, use the **[Insert](PowerPoint.FreeformBuilder.ConvertToShape.md)** method of the **[ShapeNodes](PowerPoint.ShapeNodes.md)** collection.


## Syntax

 _expression_. **AddNodes**(**_SegmentType_**, **_EditingType_**, **_X1_**, **_Y1_**, **_X2_**, **_Y2_**, **_X3_**, **_Y3_**)




## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SegmentType_|Required|**[MsoSegmentType](Office.MsoSegmentType.md)**|The type of segment to be added.|
| _EditingType_|Required|**[MsoEditingType](Office.MsoEditingType.md)**|The editing property of the vertex. If SegmentType is  **msoSegmentLine**, EditingType must be **msoEditingAuto**.|
| _X1_|Required|**Single**|If the EditingType of the new segment is  **msoEditingAuto**, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the endpoint of the new segment. If the EditingType of the new node is **msoEditingCorner**, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the first control point for the new segment.|
| _Y1_|Required|**Single**|If the EditingType of the new segment is  **msoEditingAuto**, this argument specifies the vertical distance (in points) from the upper-left corner of the document to the endpoint of the new segment. If the EditingType of the new node is **msoEditingCorner**, this argument specifies the vertical distance (in points) from the upper-left corner of the document to the first control point for the new segment.|
| _X2_|Optional|**Single**|If the EditingType of the new segment is  **msoEditingCorner**, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the second control point for the new segment. If the EditingType of the new segment is **msoEditingAuto**, don't specify a value for this argument.|
| _Y2_|Optional|**Single**|If the EditingType of the new segment is  **msoEditingCorner**, this argument specifies the vertical distance (in points) from the upper-left corner of the document to the second control point for the new segment. If the EditingType of the new segment is **msoEditingAuto**, don't specify a value for this argument.|
| _X3_|Optional|**Single**|If the EditingType of the new segment is  **msoEditingCorner**, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the endpoint of the new segment. If the EditingType of the new segment is **msoEditingAuto**, don't specify a value for this argument.|
| _Y3_|Optional|**Single**|If the EditingType of the new segment is  **msoEditingCorner**, this argument specifies the vertical distance (in points) from the upper-left corner of the document to the endpoint of the new segment. If the EditingType of the new segment is **msoEditingAuto**, don't specify a value for this argument.|

## Example

This example adds a freeform with five vertices to the first slide in the active presentation.


```vb
Set myDocument = ActivePresentation.Slides(1) 
With myDocument.Shapes.BuildFreeform(msoEditingCorner, 360, 200) 
    .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingCorner, _ 
        X1:=380, Y1:=230, X2:=400, Y2:=250, X3:=450, Y3:=300 
    .AddNodes SegmentType:=msoSegmentCurve, EditingType:=msoEditingAuto, _ 
        X1:=480, Y1:=200 
    .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _ 
        X1:=480, Y1:=400 
    .AddNodes SegmentType:=msoSegmentLine, EditingType:=msoEditingAuto, _ 
        X1:=360, Y1:=200 
    .ConvertToShape 
End With
```


## See also


[FreeformBuilder Object](PowerPoint.FreeformBuilder.md)

