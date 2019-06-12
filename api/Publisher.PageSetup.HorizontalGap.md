---
title: PageSetup.HorizontalGap property (Publisher)
keywords: vbapb10.chm6946818
f1_keywords:
- vbapb10.chm6946818
ms.prod: publisher
api_name:
- Publisher.PageSetup.HorizontalGap
ms.assetid: e8ee51e0-59b3-8fb6-21f6-87d67a96dd66
ms.date: 06/12/2019
localization_priority: Normal
---


# PageSetup.HorizontalGap property (Publisher)

Returns a **Variant** that represents the distance between the right edge of one publication page and the left edge of the next publication page in the same row when multiple pages are printed on one sheet of printer paper. Read-only.


## Syntax

_expression_.**HorizontalGap**

_expression_ A variable that represents a **[PageSetup](Publisher.PageSetup.md)** object.


## Return value

Variant


## Remarks

Numeric values are evaluated as [points](../language/glossary/vbe-glossary.md#point); string values can be in any unit supported by Microsoft Publisher (for example, "2.5 in"). The valid range of possible values is from zero to the difference between the sheet width and the page width.

This property applies only to publications where multiple pages are printed on each printer sheet. Using this property for any other publication raises an error.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]