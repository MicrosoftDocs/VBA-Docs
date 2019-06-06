---
title: Document.DefaultTabStop property (Publisher)
keywords: vbapb10.chm196616
f1_keywords:
- vbapb10.chm196616
ms.prod: publisher
api_name:
- Publisher.Document.DefaultTabStop
ms.assetid: 245ff7a3-9828-5220-b692-2ce6effb9eb6
ms.date: 06/06/2019
localization_priority: Normal
---


# Document.DefaultTabStop property (Publisher)

Returns or sets a **Variant** corresponding to the default tab stop for all text in the active publication. Valid range is 1 to 1584 [points](../language/glossary/vbe-glossary.md#point) (0.014" to 22"). Once set, numeric values are considered to be in points. **String** values may be in any unit supported by Microsoft Publisher. Point values are always returned. If values are outside the valid range, an error is returned. Read/write.


## Syntax

_expression_.**DefaultTabStop**

_expression_ A variable that represents a **[Document](Publisher.Document.md)** object.


## Return value

Variant


## Remarks

Use the **[InchesToPoints](Publisher.Application.InchesToPoints.md)** method to convert inches to points.


## Example

This example sets the **DefaultTabStop** property to 72 points for all text in the active publication.

```vb
Sub SetTab() 
 Application.ActiveDocument.DefaultTabStop = 72 
End Sub 
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]