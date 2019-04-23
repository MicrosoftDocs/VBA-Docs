---
title: Shapes.BuildFreeform method (PowerPoint)
keywords: vbapp10.chm543015
f1_keywords:
- vbapp10.chm543015
ms.prod: powerpoint
api_name:
- PowerPoint.Shapes.BuildFreeform
ms.assetid: 330ea348-9f8c-c418-d67f-e4fd6c105c59
ms.date: 09/17/2018
localization_priority: Normal
---


# Shapes.BuildFreeform method (PowerPoint)

Builds a freeform object. Returns a **[FreeformBuilder](PowerPoint.FreeformBuilder.md)** object that represents the freeform as it is being built.


## Syntax

_expression_. `BuildFreeform`( `_EditingType_`, `_X1_`, `_Y1_` )

_expression_ A variable that represents a [Shapes](PowerPoint.Shapes.md) object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _EditingType_|Required|**[MsoEditingType](Office.MsoEditingType.md)**|The editing property of the first node.|
| _X1_|Required|**Single**|The horizontal position, measured in points, of the first node in the freeform drawing relative to the left edge of the slide.|
| _Y1_|Required|**Single**|The vertical position, measured in points, of the first node in the freeform drawing relative to the top edge of the slide.|

## Return value

FreeformBuilder


## Remarks

Use the **[AddNodes](PowerPoint.FreeformBuilder.AddNodes.md)** method to add segments to the freeform. After you have added at least one segment to the freeform, you can use the **[ConvertToShape](PowerPoint.FreeformBuilder.ConvertToShape.md)** method to convert the **FreeformBuilder** object into a **[Shape](PowerPoint.Shape.md)** object that has the geometric description you've defined in the **FreeformBuilder** object.


## Example

This example adds a freeform with four segments to _myDocument_.


```vb
    Set myDocument = ActivePresentation.Slides(1)
    With myDocument.Shapes.BuildFreeform(EditingType:=msoEditingCorner, _
    X1:=360, Y1:=200) 
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

- [Shapes object](PowerPoint.Shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]