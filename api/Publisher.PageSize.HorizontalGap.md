---
title: PageSize.HorizontalGap property (Publisher)
keywords: vbapb10.chm8847368
f1_keywords:
- vbapb10.chm8847368
ms.prod: publisher
api_name:
- Publisher.PageSize.HorizontalGap
ms.assetid: 14c14534-c1c7-db2d-c7bf-8b7fd66c245e
ms.date: 06/12/2019
localization_priority: Normal
---


# PageSize.HorizontalGap property (Publisher)

Returns a **Variant** that represents the distance between the right edge of one publication page and the left edge of the next publication page in the same row in the blank page size represented by the parent **PageSize** object when multiple pages are printed on one sheet of printer paper. Read-only.


## Syntax

_expression_.**HorizontalGap**

_expression_ A variable that represents a **[PageSize](Publisher.PageSize.md)** object.


## Return value

Variant


## Remarks

The blank page size represented by the parent **PageSize** object corresponds to one of the icons displayed under **Blank Page Sizes** in the **Page Setup** dialog box in the Microsoft Publisher user interface.

Numeric values are evaluated as [points](../language/glossary/vbe-glossary.md#point); string values can be in any unit supported by Microsoft Publisher (for example, "2.5 in"). The valid range of possible values is from zero to the difference between the sheet width and the page width.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]