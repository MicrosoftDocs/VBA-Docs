---
title: FreeformBuilder.ConvertToShape method (Excel)
keywords: vbaxl10.chm648074
f1_keywords:
- vbaxl10.chm648074
ms.prod: excel
api_name:
- Excel.FreeformBuilder.ConvertToShape
ms.assetid: 2084277d-7e6a-5675-8e46-17522c3228eb
ms.date: 04/26/2019
localization_priority: Normal
---


# FreeformBuilder.ConvertToShape method (Excel)

Creates a shape that has the geometric characteristics of the specified **[FreeformBuilder](Excel.FreeformBuilder.md)** object. Returns a **[Shape](Excel.Shape.md)** object that represents the new shape.


## Syntax

_expression_.**ConvertToShape**

_expression_ A variable that represents a **[FreeformBuilder](Excel.FreeformBuilder.md)** object.


## Return value

Shape


## Remarks

You must apply the **[AddNodes](Excel.FreeformBuilder.AddNodes.md)** method to a **FreeformBuilder** object at least once before you use the **ConvertToShape** method.


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