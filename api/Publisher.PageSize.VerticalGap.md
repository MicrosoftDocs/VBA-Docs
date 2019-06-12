---
title: PageSize.VerticalGap property (Publisher)
keywords: vbapb10.chm8847369
f1_keywords:
- vbapb10.chm8847369
ms.prod: publisher
api_name:
- Publisher.PageSize.VerticalGap
ms.assetid: cc6e66ff-9a74-d88f-cfde-2f5bee66432f
ms.date: 06/12/2019
localization_priority: Normal
---


# PageSize.VerticalGap property (Publisher)

Returns a **Variant** that represents the distance in [points](../language/glossary/vbe-glossary.md#point) between the bottom edge of one publication page and the top edge of the publication page immediately below it for the blank page size represented by the parent **PageSize** object. This property applies only when multiple pages are printed on a single sheet of printer paper. Read-only.


## Syntax

_expression_.**VerticalGap**

_expression_ A variable that represents a **[PageSize](Publisher.PageSize.md)** object.


## Return value

Variant


## Remarks

The blank page size represented by the parent **PageSize** object corresponds to one of the icons displayed under **Blank Page Sizes** in the **Page Setup** dialog box in the Microsoft Publisher user interface.

Numeric values are evaluated as points. String values can be in any unit supported by Microsoft Publisher (for example, "2.5 in"). The valid range of possible values is from zero to the difference between the sheet height and the page height.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]