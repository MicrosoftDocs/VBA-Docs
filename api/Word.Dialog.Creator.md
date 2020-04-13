---
title: Dialog.Creator property (Word)
keywords: vbawd10.chm163085572
f1_keywords:
- vbawd10.chm163085572
ms.prod: word
api_name:
- Word.Dialog.Creator
ms.assetid: 73551ae3-a17d-4354-8bea-9166c5e16928
ms.date: 06/08/2017
localization_priority: Normal
---


# Dialog.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a '[Dialog](Word.Dialog.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[Dialog Object](Word.Dialog.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]