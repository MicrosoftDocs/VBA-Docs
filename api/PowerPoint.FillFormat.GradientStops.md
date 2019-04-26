---
title: FillFormat.GradientStops property (PowerPoint)
keywords: vbapp10.chm552025
f1_keywords:
- vbapp10.chm552025
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.GradientStops
ms.assetid: dd0c2c5a-81f1-b008-5b2f-5248241ac0db
ms.date: 06/08/2017
localization_priority: Normal
---


# FillFormat.GradientStops property (PowerPoint)

 Returns the **[GradientStops](Office.GradientStops.md)** collection associated with the specified fill format. Read-only.


## Syntax

_expression_.**GradientStops**

_expression_ An expression that returns a **[FillFormat](powerpoint.fillformat.md)** object.


## Return value

GradientStops


## Remarks

You can use the  **[GradientStops.Insert](Office.GradientStops.Insert.md)** method to add gradient stops to the **GradientStops** collection for the specified object.


## Example

The following example shows how to add a gradient stop at the 50% position to the  **GradientStops** collection of the fill format of the first shape on the first slide of the active presentation. For this example to work, the shape must already have a gradient fill applied.


```vb
Public Sub GradientStops_Example() 
 
    Dim pptShape As Shape 
    Dim pptFillFormat As FillFormat 
    Dim pptGradientStops As GradientStops 
     
    Set pptShape = ActivePresentation.Slides(1).Shapes(1) 
    Set pptFillFormat = pptShape.Fill 
    Set pptGradientStops = pptFillFormat.GradientStops 
     
    pptGradientStops.Insert RGB(255, 0, 255), 0.5 
     
End Sub
```


## See also


[FillFormat Object](PowerPoint.FillFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]