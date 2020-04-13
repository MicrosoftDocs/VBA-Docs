---
title: PageNumber.Creator property (Word)
ms.prod: word
api_name:
- Word.PageNumber.Creator
ms.assetid: f83e5112-c0f4-523c-e6ed-43aa572c3e2c
ms.date: 06/08/2017
localization_priority: Normal
---


# PageNumber.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a **[PageNumber](Word.PageNumber.md)** object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[PageNumber Object](Word.PageNumber.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]