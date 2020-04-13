---
title: Interior.Creator property (Word)
keywords: vbawd10.chm2818197
f1_keywords:
- vbawd10.chm2818197
ms.prod: word
api_name:
- Word.Interior.Creator
ms.assetid: 1a086cf3-9fe4-8ef4-4d9d-77d99fa8d4e2
ms.date: 06/08/2017
localization_priority: Normal
---


# Interior.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents an '[Interior](Word.Interior.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD". This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


[Interior Object](Word.Interior.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]