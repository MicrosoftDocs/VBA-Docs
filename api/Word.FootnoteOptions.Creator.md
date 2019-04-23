---
title: FootnoteOptions.Creator property (Word)
keywords: vbawd10.chm170132457
f1_keywords:
- vbawd10.chm170132457
ms.prod: word
api_name:
- Word.FootnoteOptions.Creator
ms.assetid: d63db4c6-481e-643f-0377-a33b564284bd
ms.date: 06/08/2017
localization_priority: Normal
---


# FootnoteOptions.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a '[FootnoteOptions](Word.FootnoteOptions.md)' collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[FootnoteOptions Object](Word.FootnoteOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]