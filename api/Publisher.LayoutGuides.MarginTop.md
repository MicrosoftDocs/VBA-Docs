---
title: LayoutGuides.MarginTop property (Publisher)
keywords: vbapb10.chm1114118
f1_keywords:
- vbapb10.chm1114118
ms.prod: publisher
api_name:
- Publisher.LayoutGuides.MarginTop
ms.assetid: f0b4f600-6c79-060b-edd5-82f07f78770a
ms.date: 06/08/2019
localization_priority: Normal
---


# LayoutGuides.MarginTop property (Publisher)

Returns or sets a **Variant** that represents the amount of space (in [points](../language/glossary/vbe-glossary.md#point)) between the text and the top edge of a cell, text frame, or page. Read/write.


## Syntax

_expression_.**MarginTop**

_expression_ A variable that represents a **[LayoutGuides](Publisher.LayoutGuides.md)** object.


## Example

This example sets the margins of the active publication to two inches.

```vb
Sub SetPageMargins() 
 
 With ActiveDocument.LayoutGuides 
 .MarginTop = Application.InchesToPoints(Value:=2) 
 .MarginBottom = Application.InchesToPoints(Value:=2) 
 .MarginLeft = Application.InchesToPoints(Value:=2) 
 .MarginRight = Application.InchesToPoints(Value:=2) 
 End With 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]