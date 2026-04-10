---
title: Shape.Glow property (PowerPoint)
keywords: vbapp10.chm547082
f1_keywords:
- vbapp10.chm547082
api_name:
- PowerPoint.Shape.Glow
ms.assetid: 58bea564-b90a-4b39-53c7-8bb6f6ccd019
ms.date: 06/08/2017
ms.localizationpriority: medium
---


# Shape.Glow property (PowerPoint)

Returns a  **[GlowFormat](Office.GlowFormat.md)** object that contains glow formatting properties for the specified shape. Read-only.


## Syntax

_expression_.**Glow**

_expression_ A variable that represents a **[Shape](PowerPoint.Shape.md)** object.


## Return value

GlowFormat


## Example

This example sets the color, radius, and transparency for the glow of the second shape on the second slide in a PowerPoint presentation:


```
With ActivePresentation.Slides(2).Shapes(2).Glow
    .Color.RGB = RGB(128, 0, 0)
    .Radius = 10
    .Transparency = 0.5
End With 

```


## See also


[Shape Object](PowerPoint.Shape.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
