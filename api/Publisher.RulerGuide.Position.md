---
title: RulerGuide.Position property (Publisher)
keywords: vbapb10.chm655364
f1_keywords:
- vbapb10.chm655364
ms.prod: publisher
api_name:
- Publisher.RulerGuide.Position
ms.assetid: af169eaf-3cca-0310-c49b-369ba9b2193f
ms.date: 06/13/2019
localization_priority: Normal
---


# RulerGuide.Position property (Publisher)

Returns or sets a **Variant** representing the font position relative to the baseline of the text in the specified range. Positive values move the text above the normal baseline; negative values move the text below the baseline. Indeterminate values are returned as -9999.0. Read/write.


## Syntax

_expression_.**Position**

_expression_ A variable that represents a **[RulerGuide](Publisher.RulerGuide.md)** object.


## Example

This example adjusts the text in the second story to 5 points below the baseline.

```vb
Sub Position() 
 
 Application.ActiveDocument.Stories(2).TextRange.Font.Position = -5 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]