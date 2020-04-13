---
title: Envelope.Creator property (Word)
keywords: vbawd10.chm152568809
f1_keywords:
- vbawd10.chm152568809
ms.prod: word
api_name:
- Word.Envelope.Creator
ms.assetid: bb631423-89b4-cf3e-55a9-562b8b6aaad0
ms.date: 06/08/2017
localization_priority: Normal
---


# Envelope.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents an '[Envelope](Word.Envelope.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[Envelope Object](Word.Envelope.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]