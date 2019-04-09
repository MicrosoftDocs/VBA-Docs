---
title: ThreeDFormat.Creator property (Word)
keywords: vbawd10.chm164627433
f1_keywords:
- vbawd10.chm164627433
ms.prod: word
api_name:
- Word.ThreeDFormat.Creator
ms.assetid: 91566d7d-5934-f0a2-8bfc-16054b5c9292
ms.date: 06/08/2017
localization_priority: Normal
---


# ThreeDFormat.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a '[ThreeDFormat](Word.ThreeDFormat.md)' object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[ThreeDFormat Object](Word.ThreeDFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]