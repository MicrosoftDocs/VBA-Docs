---
title: Shape.Left property (Publisher)
keywords: vbapb10.chm2228289
f1_keywords:
- vbapb10.chm2228289
ms.prod: publisher
api_name:
- Publisher.Shape.Left
ms.assetid: 275f5af9-9812-2a6b-bba3-704d4a7f5601
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.Left property (Publisher)

Returns or sets a **Variant** indicating the distance from the left edge of the page to the leftmost edge of the specified shape. Numeric values are in [points](../language/glossary/vbe-glossary.md#point); all other values are in any measurement supported by Publisher (for example, "2.5 in"). Read/write.


## Syntax

_expression_.**Left**

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Example

This example sets the horizontal position of the first shape in the active publication to 1 inch from the left edge of the page.

```vb
With ActiveDocument.Pages(1).Shapes(1) 
 .Left = InchesToPoints(1) 
End With
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]