---
title: Paragraphs.Creator property (Word)
keywords: vbawd10.chm156763113
f1_keywords:
- vbawd10.chm156763113
ms.prod: word
api_name:
- Word.Paragraphs.Creator
ms.assetid: f858f81f-3e41-77a9-9a98-d7dd60fa2e0a
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a '[Paragraphs](Word.paragraphs.md)' collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[Paragraphs Collection Object](Word.paragraphs.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]