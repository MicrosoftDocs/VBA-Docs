---
title: Shapes.BuildFreeform method (Excel)
keywords: vbaxl10.chm638087
f1_keywords:
- vbaxl10.chm638087
ms.prod: excel
api_name:
- Excel.Shapes.BuildFreeform
ms.assetid: 0eec4b87-1a40-1e60-a66a-a8bb2b2f7efa
ms.date: 05/15/2019
localization_priority: Normal
---


# Shapes.BuildFreeform method (Excel)

Builds a freeform object. Returns a **[FreeformBuilder](Excel.FreeformBuilder.md)** object that represents the freeform as it is being built. 

Use the **[AddNodes](Excel.FreeformBuilder.AddNodes.md)** method to add segments to the freeform. After you have added at least one segment to the freeform, you can use the **[ConvertToShape](Excel.FreeformBuilder.ConvertToShape.md)** method to convert the **FreeformBuilder** object into a **[Shape](Excel.Shape.md)** object that has the geometric description that you have defined in the **FreeformBuilder** object.


## Syntax

_expression_.**BuildFreeform** (_EditingType_, _X1_, _Y1_)

_expression_ A variable that represents a **[Shapes](Excel.Shapes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _EditingType_|Required| **[MsoEditingType](Office.MsoEditingType.md)**|The editing property of the first node.|
| _X1_|Required| **Single**|The position (in [points](../language/glossary/vbe-glossary.md#point)) of the first node in the freeform drawing relative to the upper-left corner of the document.|
| _Y1_|Required| **Single**|The position (in points) of the first node in the freeform drawing relative to the upper-left corner of the document.|

## Return value

**FreeformBuilder**


## Example

This example adds a freeform with five vertices to _myDocument_.

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