---
title: FreeformBuilder.ConvertToShape method (Word)
keywords: vbawd10.chm164167691
f1_keywords:
- vbawd10.chm164167691
ms.prod: word
api_name:
- Word.FreeformBuilder.ConvertToShape
ms.assetid: 9c88065f-1ff0-ac53-a630-4f0b4e652a80
ms.date: 06/08/2017
localization_priority: Normal
---


# FreeformBuilder.ConvertToShape method (Word)

Creates a shape that has the geometric characteristics of the specified object. Returns a  **[Shape](Word.Shape.md)** object that represents the new shape.


## Syntax

_expression_. `ConvertToShape`( `_Anchor_` )

_expression_ Required. A variable that represents a '[FreeformBuilder](Word.FreeformBuilder.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Anchor_|Optional| **Variant**|A **[Range](Word.Range.md)** object that represents the text to which the shape is bound. If Anchor is specified, the anchor is positioned at the beginning of the first paragraph in the anchoring range. If this argument is omitted, the anchoring range is selected automatically and the shape is positioned relative to the top and left edges of the page.|

## Remarks

You must apply the **[AddNodes](Word.FreeformBuilder.AddNodes.md)** method to a **FreeformBuilder** object at least once before you use the **ConvertToShape** method.


## See also


[FreeformBuilder Object](Word.FreeformBuilder.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]