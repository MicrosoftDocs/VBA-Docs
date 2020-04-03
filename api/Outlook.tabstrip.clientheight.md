---
title: TabStrip.ClientHeight Property (Outlook Forms Script)
keywords: olfm10.chm2000890
f1_keywords:
- olfm10.chm2000890
ms.prod: outlook
ms.assetid: 937ca019-5d32-bb82-8359-a74e4da12c9f
ms.date: 06/08/2017
localization_priority: Normal
---


# TabStrip.ClientHeight Property (Outlook Forms Script)

Returns a  **Single** value that represents the height dimension of the display area of a **[TabStrip](Outlook.tabstrip.md)**. Read-only.


## Syntax

_expression_.**ClientHeight**

_expression_ A variable that represents a  **TabStrip** object.


## Remarks

For  **ClientHeight** and **[ClientWidth](Outlook.tabstrip.clientwidth.md)**, specifies the distance, in [points](../language/glossary/vbe-glossary.md#point), from respectively the top and left edge of the TabStrip's container. For  **[ClientLeft](Outlook.tabstrip.clientleft.md)** and **[ClientTop](Outlook.tabstrip.clienttop.md)**, specifies the location, in [points](../language/glossary/vbe-glossary.md#point), of respectively the top and left edges of the TabStrip's container.

At run time,  **ClientLeft**,  **ClientTop**,  **ClientHeight**, and  **ClientWidth** automatically store the coordinates and dimensions of the **TabStrip** control's internal area, which is shared by objects in the **TabStrip**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]