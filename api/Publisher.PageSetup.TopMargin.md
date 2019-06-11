---
title: PageSetup.TopMargin property (Publisher)
keywords: vbapb10.chm6946837
f1_keywords:
- vbapb10.chm6946837
ms.prod: publisher
api_name:
- Publisher.PageSetup.TopMargin
ms.assetid: 4eee9b1e-6c76-7831-13bc-25926c3c0f10
ms.date: 06/12/2019
localization_priority: Normal
---


# PageSetup.TopMargin property (Publisher)

Returns a **Variant** that represents the distance between the top edge of the printer sheet and the top edge of the publication pages. Read-only.


## Syntax

_expression_.**TopMargin**

_expression_ A variable that represents a **[PageSetup](Publisher.PageSetup.md)** object.


## Return value

Variant


## Remarks

Numeric values are evaluated as points. String values can be in any unit supported by Microsoft Publisher (for example, "2.5 in"). The valid range of possible values is from zero to the difference between the height of the sheet and the height of the publication pages.

The **TopMargin** property returns a value only when you print multiple pages on a single sheet of printer paper. If you attempt to use it in other circumstances, Microsoft Publisher returns **Nothing**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]