---
title: FreeformBuilder.AddNodes method (Excel)
keywords: vbaxl10.chm648073
f1_keywords:
- vbaxl10.chm648073
ms.prod: excel
api_name:
- Excel.FreeformBuilder.AddNodes
ms.assetid: 8fff188d-1c47-87f0-8388-2b12534e82c2
ms.date: 04/26/2019
localization_priority: Normal
---


# FreeformBuilder.AddNodes method (Excel)

Adds a point in the current shape, and then draws a line from the current node to the last node that was added.


## Syntax

_expression_.**AddNodes** (_SegmentType_, _EditingType_, _X1_, _Y1_, _X2_, _Y2_, _X3_, _Y3_)

_expression_ A variable that represents a **[FreeformBuilder](Excel.FreeformBuilder.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SegmentType_|Required| **[MsoSegmentType](Office.MsoSegmentType.md)**|The type of segment to be added.|
| _EditingType_|Required| **[MsoEditingType](Office.MsoEditingType.md)**|The editing property of the vertex.|
| _X1_|Required| **Single**|If the _EditingType_ of the new segment is **msoEditingAuto**, this argument specifies the horizontal distance (in [points](../language/glossary/vbe-glossary.md#point)) from the upper-left corner of the document to the end point of the new segment.<br/><br/>If the _EditingType_ of the new node is **msoEditingCorner**, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the first control point for the new segment.|
| _Y1_|Required| **Single**|If the _EditingType_ of the new segment is **msoEditingAuto**, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the end point of the new segment.<br/><br/>If the _EditingType_ of the new node is **msoEditingCorner**, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the first control point for the new segment.|
| _X2_|Optional| **Variant**|If the _EditingType_ of the new segment is **msoEditingCorner**, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the second control point for the new segment.<br/><br/>If the _EditingType_ of the new segment is **msoEditingAuto**, don't specify a value for this argument.|
| _Y2_|Optional| **Variant**|If the _EditingType_ of the new segment is **msoEditingCorner**, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the second control point for the new segment.<br/><br/>If the _EditingType_ of the new segment is **msoEditingAuto**, don't specify a value for this argument.|
| _X3_|Optional| **Variant**|If the _EditingType_ of the new segment is **msoEditingCorner**, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the second control point for the new segment.<br/><br/>If the _EditingType_ of the new segment is **msoEditingAuto**, don't specify a value for this argument.|
| _Y3_|Optional| **Variant**|If the _EditingType_ of the new segment is **msoEditingCorner**, this argument specifies the horizontal distance (in points) from the upper-left corner of the document to the second control point for the new segment.<br/><br/>If the _EditingType_ of the new segment is **msoEditingAuto**, don't specify a value for this argument.|

## Remarks

**MsoEditingType** cannot be **msoEditingSmooth** or **msoEditingSymmetric**. If _SegmentType_ is **msoSegmentLine**, _EditingType_ must be **msoEditingAuto**.

## Example

This example adds a freeform with four segments to _myDocument_.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes.BuildFreeform(msoEditingCorner, 360, 200) 
    .AddNodes msoSegmentCurve, msoEditingCorner, _ 
        380, 230, 400, 250, 450, 300 
    .AddNodes msoSegmentCurve, msoEditingAuto, 480, 200 
    .AddNodes msoSegmentLine, msoEditingAuto, 480, 400 
    .AddNodes msoSegmentLine, msoEditingAuto, 360, 200 
    .ConvertToShape 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]