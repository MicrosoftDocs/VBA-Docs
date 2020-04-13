---
title: Range.Creator property (Word)
keywords: vbawd10.chm157156329
f1_keywords:
- vbawd10.chm157156329
ms.prod: word
api_name:
- Word.Range.Creator
ms.assetid: df19ebef-edb2-ac75-e878-bce9e35794b3
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a **[Range](Word.Range.md)** object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]