---
title: Page.ScrollLeft Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 8a8be730-5dca-5ad7-2f08-370fc0a95dd3
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.ScrollLeft Property (Outlook Forms Script)

Returns or sets a  **Single** that specifies the distance, in [points](../language/glossary/vbe-glossary.md#point), of the left edge of the visible form from the left edge of the **[Page](Outlook.page.md)**. Read/write.


## Syntax

_expression_.**ScrollLeft**

_expression_ A variable that represents a  **Page** object.


## Remarks

The minimum value is zero; the maximum value is the difference between the value of the  **[ScrollWidth](Outlook.page.scrollwidth.md)** property and the value of the **Width** property for the form or page.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]