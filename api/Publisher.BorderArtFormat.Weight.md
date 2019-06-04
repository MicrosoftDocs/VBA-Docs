---
title: BorderArtFormat.Weight property (Publisher)
keywords: vbapb10.chm7602182
f1_keywords:
- vbapb10.chm7602182
ms.prod: publisher
api_name:
- Publisher.BorderArtFormat.Weight
ms.assetid: 8ff67c8b-be41-a02e-5433-624baa0d888e
ms.date: 06/05/2019
localization_priority: Normal
---


# BorderArtFormat.Weight property (Publisher)

Returns or sets a **Variant** indicating the thickness of the specified line or cell border.


## Syntax

_expression_.**Weight**

_expression_ A variable that represents a **[BorderArtFormat](Publisher.BorderArtFormat.md)** object.


## Remarks

Return values are in [points](../language/glossary/vbe-glossary.md#point). When setting the property, numeric values are evaluated in points, and strings can be in any units supported by Publisher (for example, "2.5 in").


## Example

This example adds a green dashed line, two points thick, to the active publication.

```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddLine(BeginX:=10, BeginY:=10, _ 
 EndX:=250, EndY:=250).Line 
 .DashStyle = msoLineDashDotDot 
 .ForeColor.RGB = RGB(0, 255, 255) 
 .Weight = 2 
End With 

```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]