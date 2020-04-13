---
title: Languages.Creator property (Word)
keywords: vbawd10.chm158204905
f1_keywords:
- vbawd10.chm158204905
ms.prod: word
api_name:
- Word.Languages.Creator
ms.assetid: ef63da8d-146f-9e62-3179-dc070580a68a
ms.date: 06/08/2017
localization_priority: Normal
---


# Languages.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a '[Languages](Word.languages.md)' collection.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[Languages Collection Object](Word.languages.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]