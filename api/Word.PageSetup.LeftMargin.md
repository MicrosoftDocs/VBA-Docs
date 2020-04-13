---
title: PageSetup.LeftMargin property (Word)
keywords: vbawd10.chm158400614
f1_keywords:
- vbawd10.chm158400614
ms.prod: word
api_name:
- Word.PageSetup.LeftMargin
ms.assetid: 873d6cf2-da9f-5d88-314f-9820284a54ee
ms.date: 06/08/2017
localization_priority: Normal
---


# PageSetup.LeftMargin property (Word)

Returns or sets the distance (in points) between the left edge of the page and the left boundary of the body text. Read/write  **Single**.


## Syntax

_expression_.**LeftMargin**

 _expression_ An expression that returns a **[PageSetup](Word.PageSetup.md)** object.


## Remarks

If the **[MirrorMargins](Word.PageSetup.MirrorMargins.md)** property is set to **True**, the LeftMargin property controls the setting for inside margins and the **[RightMargin](Word.PageSetup.RightMargin.md)** property controls the setting for outside margins.


## Example

This example sets the left margin to 1 inch (72 points) for the second section in the active document.


```vb
ActiveDocument.Sections(2).PageSetup.LeftMargin = 72
```


## See also


[PageSetup Object](Word.PageSetup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]