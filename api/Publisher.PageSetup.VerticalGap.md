---
title: PageSetup.VerticalGap property (Publisher)
keywords: vbapb10.chm6946838
f1_keywords:
- vbapb10.chm6946838
ms.prod: publisher
api_name:
- Publisher.PageSetup.VerticalGap
ms.assetid: 191d66c4-d168-625a-47b7-028167a98af9
ms.date: 06/12/2019
localization_priority: Normal
---


# PageSetup.VerticalGap property (Publisher)

Returns a **Variant** that represents the distance (in [points](../language/glossary/vbe-glossary.md#point)) between the bottom edge of one publication page and the top edge of the publication page below it when more than one publication page is printed on a single printer page. Read-only.


## Syntax

_expression_.**VerticalGap**

_expression_ A variable that represents a **[PageSetup](Publisher.PageSetup.md)** object.


## Return value

Variant


## Remarks

You can use the **VerticalGap** property when you want to print multiple pages on a single sheet of printer paper. If the page size, including the values for the **VerticalGap** and **HorizontalGap** properties, is greater than half the paper size, Microsoft Publisher displays an error.


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]