---
title: GradientStop object (Office)
ms.prod: office
api_name:
- Office.GradientStop
ms.assetid: b5003bfc-9ac6-fd56-f214-a0d99db0cf07
ms.date: 01/16/2019
localization_priority: Normal
---


# GradientStop object (Office)

Represents one gradient stop.


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

- [GradientStop object members](overview/library-reference/gradientstop-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]