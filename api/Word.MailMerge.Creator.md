---
title: MailMerge.Creator property (Word)
keywords: vbawd10.chm153093097
f1_keywords:
- vbawd10.chm153093097
ms.prod: word
api_name:
- Word.MailMerge.Creator
ms.assetid: b72970e2-160d-3b8d-4ada-a78957ff1e73
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMerge.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a '[MailMerge](Word.MailMerge.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[MailMerge Object](Word.MailMerge.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]