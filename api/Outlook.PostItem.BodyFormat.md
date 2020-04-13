---
title: PostItem.BodyFormat property (Outlook)
keywords: vbaol11.chm1558
f1_keywords:
- vbaol11.chm1558
ms.prod: outlook
api_name:
- Outlook.PostItem.BodyFormat
ms.assetid: 4d60e71c-d492-5ba4-b9d2-e61fb608abcc
ms.date: 06/08/2017
localization_priority: Normal
---


# PostItem.BodyFormat property (Outlook)

Returns or sets an **[OlBodyFormat](Outlook.OlBodyFormat.md)** constant indicating the format of the body text. Read/write.


## Syntax

_expression_. `BodyFormat`

_expression_ A variable that represents a [PostItem](Outlook.PostItem.md) object.


## Remarks

The body text format determines the standard used to display the text of the message. Microsoft Outlook provides three body text format options: Plain Text, Rich Text (RTF), and HTML.

All text formatting will be lost when the  **BodyFormat** property is switched from RTF to HTML and vice-versa.


## See also


[PostItem Object](Outlook.PostItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]