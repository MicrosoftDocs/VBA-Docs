---
title: GradientStops.Delete method (Office)
ms.prod: office
api_name:
- Office.GradientStops.Delete
ms.assetid: 3f31656a-498d-57d1-1464-b2439718ef89
ms.date: 01/16/2019
localization_priority: Normal
---


# GradientStops.Delete method (Office)

Removes a gradient stop.


## Syntax

_expression_.**Delete** (_Index_)

_expression_ An expression that returns a **[GradientStops](Office.GradientStops.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Optional|**Integer**|The index number of the gradient stop.|

## Remarks

Gradients are a smooth transition from one color state to another. The endpoints of these sections are called stops.


## Example

The following example adds three gradient color stops and then deletes the first gradient stop.


```vb
Sub gradients() 
 Set myDocument = ActivePresentation.Slides(1) 
 Set GradientShapeFill = myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 80).Fill 
 With GradientShapeFill 
 .ForeColor.RGB = RGB(0, 128, 128) 
 .OneColorGradient msoGradientHorizontal, 1, 1 
 .GradientStops.Insert RGB(255, 0, 0), 0.25 
 .GradientStops.Insert RGB(0, 255, 0), 0.5 
 .GradientStops.Insert RGB(0, 0, 255), 0.75 
 End With 
 GradientShapeFill.GradientStops.Delete (1) 
End Sub
```


## See also

- [GradientStops object members](overview/library-reference/gradientstops-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]