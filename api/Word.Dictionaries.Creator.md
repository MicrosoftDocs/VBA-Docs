---
title: Dictionaries.Creator property (Word)
keywords: vbawd10.chm162268137
f1_keywords:
- vbawd10.chm162268137
ms.prod: word
api_name:
- Word.Dictionaries.Creator
ms.assetid: 02d51be5-8640-f441-d52c-3f2ff73e6495
ms.date: 06/08/2017
localization_priority: Normal
---


# Dictionaries.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a '[Dictionaries](Word.dictionaries.md)' collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[Dictionaries Collection Object](Word.dictionaries.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]