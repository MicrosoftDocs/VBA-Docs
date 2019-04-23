---
title: Frame.ScrollLeft Property (Outlook Forms Script)
keywords: olfm10.chm2001800
f1_keywords:
- olfm10.chm2001800
ms.prod: outlook
ms.assetid: 576d571d-05fa-2e1d-df7d-3bb1c606c374
ms.date: 06/08/2017
localization_priority: Normal
---


# Frame.ScrollLeft Property (Outlook Forms Script)

Returns or sets a  **Single** that specifies the distance, in [points](../language/glossary/vbe-glossary.md#point), of the left edge of the visible form from the left edge of the **[Frame](Outlook.frame.md)**. Read/write.


## Syntax

_expression_.**ScrollLeft**

_expression_ A variable that represents a  **Frame** object.


## Remarks

The minimum value is zero; the maximum value is the difference between the value of the  **[ScrollWidth](Outlook.frame.scrollwidth.md)** property and the value of the **Width** property for the form or page.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]