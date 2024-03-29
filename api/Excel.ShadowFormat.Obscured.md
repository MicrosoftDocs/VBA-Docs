---
title: ShadowFormat.Obscured property (Excel)
keywords: vbaxl10.chm114003
f1_keywords:
- vbaxl10.chm114003
api_name:
- Excel.ShadowFormat.Obscured
ms.assetid: a2cc3324-d394-5332-41d2-e3733d0eb2d7
ms.date: 05/14/2019
ms.localizationpriority: medium
---


# ShadowFormat.Obscured property (Excel)

**True** if the shadow of the specified shape appears filled in and is obscured by the shape, even if the shape has no fill. 

**False** if the shadow has no fill and the outline of the shadow is visible through the shape if the shape has no fill. Read/write **[MsoTriState](office.msotristate.md)**.


## Syntax

_expression_.**Obscured**

_expression_ A variable that represents a **[ShadowFormat](Excel.ShadowFormat.md)** object.



## Example

This example sets the horizontal and vertical offsets for the shadow of shape three on _myDocument_. The shadow is offset 5 [points](../language/glossary/vbe-glossary.md#point) to the right of the shape and 3 points above it. If the shape doesn't already have a shadow, this example adds one to it. The shadow will be filled in and obscured by the shape, even if the shape has no fill.

```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3).Shadow 
 .Visible = True 
 .OffsetX = 5 
 .OffsetY = -3 
 .Obscured = msoTrue 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]