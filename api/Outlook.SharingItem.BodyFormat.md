---
title: SharingItem.BodyFormat property (Outlook)
keywords: vbaol11.chm675
f1_keywords:
- vbaol11.chm675
ms.prod: outlook
api_name:
- Outlook.SharingItem.BodyFormat
ms.assetid: 60a18df9-8882-a5a2-efb9-cc59206f7345
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.BodyFormat property (Outlook)

Returns or sets an **[OlBodyFormat](Outlook.OlBodyFormat.md)** constant indicating the format of the body text. Read/write.


## Syntax

_expression_. `BodyFormat`

_expression_ A variable that represents a [SharingItem](Outlook.SharingItem.md) object.


## Remarks

The body text format determines the standard used to display the text of the message. Microsoft Outlook provides three body text format options: Plain Text, Rich Text (RTF), and HTML.

All text formatting will be lost when the  **BodyFormat** property is switched from RTF to HTML and vice-versa.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]