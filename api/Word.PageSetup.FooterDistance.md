---
title: PageSetup.FooterDistance property (Word)
keywords: vbawd10.chm158400625
f1_keywords:
- vbawd10.chm158400625
ms.prod: word
api_name:
- Word.PageSetup.FooterDistance
ms.assetid: 0c3fda7d-be19-982c-b54e-34905be189d1
ms.date: 06/08/2017
localization_priority: Normal
---


# PageSetup.FooterDistance property (Word)

Returns or sets the distance (in points) between the footer and the bottom of the page. Read/write  **Single**.


## Syntax

_expression_. `FooterDistance`

_expression_ A variable that represents a **[PageSetup](Word.PageSetup.md)** object.


## Example

This example sets the distance between the footer and the bottom of the page to 0.5 inch. The  **[InchesToPoints](Word.Application.InchesToPoints.md)** method is used to convert inches to points.


```vb
ActiveDocument.PageSetup.FooterDistance = InchesToPoints(0.5)
```

This example sets the distance between the footer and the bottom of the page for all the sections in the selection to 1 inch.




```vb
Selection.Range.PageSetup.FooterDistance = 72
```


## See also


[PageSetup Object](Word.PageSetup.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]