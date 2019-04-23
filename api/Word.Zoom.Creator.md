---
title: Zoom.Creator property (Word)
keywords: vbawd10.chm161874921
f1_keywords:
- vbawd10.chm161874921
ms.prod: word
api_name:
- Word.Zoom.Creator
ms.assetid: 6505d975-4e4d-1af3-4bc6-0b0642206550
ms.date: 06/08/2017
localization_priority: Normal
---


# Zoom.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a '[Zoom](Word.Zoom.md)' object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[Zoom Object](Word.Zoom.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]