---
title: Style.Creator property (Word)
keywords: vbawd10.chm153879529
f1_keywords:
- vbawd10.chm153879529
ms.prod: word
api_name:
- Word.Style.Creator
ms.assetid: dce5e2bb-93f9-acc4-7173-3375eaf76eff
ms.date: 06/08/2017
localization_priority: Normal
---


# Style.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a '[Style](Word.Style.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[Style Object](Word.Style.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]