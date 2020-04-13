---
title: Categories.Creator property (Word)
keywords: vbawd10.chm126551017
f1_keywords:
- vbawd10.chm126551017
ms.prod: word
api_name:
- Word.Categories.Creator
ms.assetid: 7419019c-d6b1-1f10-5bb4-7e87f8c4acf1
ms.date: 06/08/2017
localization_priority: Normal
---


# Categories.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a '[Categories](Word.Categories.md)' collection.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[Categories Object](Word.Categories.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]