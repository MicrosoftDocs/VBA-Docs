---
title: Variables.Creator property (Word)
keywords: vbawd10.chm157615081
f1_keywords:
- vbawd10.chm157615081
ms.prod: word
api_name:
- Word.Variables.Creator
ms.assetid: afef2a48-87c0-36b2-6242-31ba8f5d5d00
ms.date: 06/08/2017
localization_priority: Normal
---


# Variables.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a '[Variables](Word.variables.md)' collection.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[Variables Collection Object](Word.variables.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]