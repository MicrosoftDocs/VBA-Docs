---
title: CalloutFormat.Gap property (Publisher)
keywords: vbapb10.chm2490631
f1_keywords:
- vbapb10.chm2490631
ms.prod: publisher
api_name:
- Publisher.CalloutFormat.Gap
ms.assetid: fd7cdac7-5f09-a574-e9ef-08feebd81cff
ms.date: 06/05/2019
localization_priority: Normal
---


# CalloutFormat.Gap property (Publisher)

Returns or sets a **Variant** indicating the horizontal distance between the end of the callout line and the text bounding box. Read/write.


## Syntax

_expression_.**Gap**

_expression_ A variable that represents a **[CalloutFormat](Publisher.CalloutFormat.md)** object.


## Return value

Variant


## Remarks

Numeric values are evaluated in [points](../language/glossary/vbe-glossary.md#point); strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").


## Example

This example sets the distance between the callout line and the text bounding box to 3 points for the first shape in the active publication. For the example to work, the shape must be a callout.

```vb
ActiveDocument.Pages(1).Shapes(1).Callout.Gap = 3
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]