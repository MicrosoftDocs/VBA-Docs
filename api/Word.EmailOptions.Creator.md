---
title: EmailOptions.Creator property (Word)
keywords: vbawd10.chm165347429
f1_keywords:
- vbawd10.chm165347429
ms.prod: word
api_name:
- Word.EmailOptions.Creator
ms.assetid: 177aec79-5da3-e761-317c-32a7b2ecc23d
ms.date: 06/08/2017
localization_priority: Normal
---


# EmailOptions.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents an '[EmailOptions](Word.EmailOptions.md)' collection.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[EmailOptions Object](Word.EmailOptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]