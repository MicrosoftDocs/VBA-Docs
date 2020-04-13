---
title: Section.Creator property (Word)
keywords: vbawd10.chm156828649
f1_keywords:
- vbawd10.chm156828649
ms.prod: word
api_name:
- Word.Section.Creator
ms.assetid: 203b4c7d-29e2-cde0-b155-acd3bd68f7ff
ms.date: 06/08/2017
localization_priority: Normal
---


# Section.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a '[Section](Word.Section.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[Section Object](Word.Section.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]