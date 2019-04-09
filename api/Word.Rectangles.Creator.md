---
title: Rectangles.Creator property (Word)
ms.prod: word
api_name:
- Word.Rectangles.Creator
ms.assetid: 59f705bf-8d15-fb57-3809-3f5df35938aa
ms.date: 06/08/2017
localization_priority: Normal
---


# Rectangles.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a '[Rectangles](Word.Rectangles.md)' collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[Rectangles Collection](Word.Rectangles.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]