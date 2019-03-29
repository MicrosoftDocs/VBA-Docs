---
title: FreeformBuilder object (Excel)
keywords: vbaxl10.chm647072
f1_keywords:
- vbaxl10.chm647072
ms.prod: excel
api_name:
- Excel.FreeformBuilder
ms.assetid: 91c779ac-69bc-3b68-8ecb-1f9cc8e5b20e
ms.date: 03/30/2019
localization_priority: Normal
---


# FreeformBuilder object (Excel)

Represents the geometry of a freeform while it's being built.


## Remarks

Use the **[BuildFreeform](Excel.Shapes.BuildFreeform.md)** method of the **Shapes** object to return a **FreeformBuilder** object. Use the **AddNodes** method to add nodes to the freeform. Use the **ConvertToShape** method to create the shape defined in the **FreeformBuilder** object and add it to the **[Shapes](Excel.Shapes.md)** collection.


## Example

The following example adds a freeform with four segments to _myDocument_.


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

## Methods

- [AddNodes](Excel.FreeformBuilder.AddNodes.md)
- [ConvertToShape](Excel.FreeformBuilder.ConvertToShape.md)

## Properties

- [Application](Excel.FreeformBuilder.Application.md)
- [Creator](Excel.FreeformBuilder.Creator.md)
- [Parent](Excel.FreeformBuilder.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]