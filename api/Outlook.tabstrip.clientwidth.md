---
title: TabStrip.ClientWidth Property (Outlook Forms Script)
keywords: olfm10.chm2000920
f1_keywords:
- olfm10.chm2000920
ms.prod: outlook
ms.assetid: f59ccbe8-8f45-38d4-15f0-23fa8d52b50f
ms.date: 06/08/2017
localization_priority: Normal
---


# TabStrip.ClientWidth Property (Outlook Forms Script)

Returns a  **Single** value that represents the width dimension of the display area of a **[TabStrip](Outlook.tabstrip.md)**. Read-only.


## Syntax

_expression_.**ClientWidth**

_expression_ A variable that represents a  **TabStrip** object.


## Remarks

For  **[ClientHeight](Outlook.tabstrip.clientheight.md)** and **ClientWidth**, specifies the distance, in [points](../language/glossary/vbe-glossary.md#point), from respectively the top and left edge of the TabStrip's container. For  **[ClientLeft](Outlook.tabstrip.clientleft.md)** and **[ClientTop](Outlook.tabstrip.clienttop.md)**, specifies the location, in [points](../language/glossary/vbe-glossary.md#point), of respectively the top and left edges of the TabStrip's container.

At run time,  **ClientLeft**,  **ClientTop**,  **ClientHeight**, and  **ClientWidth** automatically store the coordinates and dimensions of the **TabStrip** control's internal area, which is shared by objects in the **TabStrip**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]