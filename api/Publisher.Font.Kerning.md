---
title: Font.Kerning property (Publisher)
keywords: vbapb10.chm5373976
f1_keywords:
- vbapb10.chm5373976
ms.prod: publisher
api_name:
- Publisher.Font.Kerning
ms.assetid: 756fe3fa-9bf3-be16-2dd1-5b8fb0ec6496
ms.date: 06/08/2019
localization_priority: Normal
---


# Font.Kerning property (Publisher)

Returns or sets a **Variant** indicating the amount of horizontal spacing that Microsoft Publisher applies to the characters in the text range. Read/write.


## Syntax

_expression_.**Kerning**

_expression_ A variable that represents a **[Font](Publisher.Font.md)** object.


## Return value

Variant


## Remarks

When setting this property, numeric values are considered to be in [points](../language/glossary/vbe-glossary.md#point), and **String** values may be in any unit supported by Publisher. Return values are of type **Single** and in points. 

Negative values bring characters closer together than normal, and positive values spread characters farther apart than normal. The valid range is -600.0 to 600.0 points.

Use the **[InchesToPoints](Publisher.Application.InchesToPoints.md)** method to convert inches to points.


## Example

This example adjusts the kerning of all text in the first story to 6 point.

```vb
Application.ActiveDocument.Stories(1).TextRange.Font.Kerning = 6 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]