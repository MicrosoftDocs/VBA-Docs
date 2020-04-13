---
title: Category.Creator property (Word)
keywords: vbawd10.chm190710761
f1_keywords:
- vbawd10.chm190710761
ms.prod: word
api_name:
- Word.Category.Creator
ms.assetid: 966613ee-09a2-3f3e-ea4b-0e6c062a5863
ms.date: 06/08/2017
localization_priority: Normal
---


# Category.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a '[Category](Word.Category.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[Category Object](Word.Category.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]