---
title: ContentControlListEntry.Creator property (Word)
keywords: vbawd10.chm147456101
f1_keywords:
- vbawd10.chm147456101
ms.prod: word
api_name:
- Word.ContentControlListEntry.Creator
ms.assetid: a16247f3-7faf-3ff5-e5c0-53d176c79ea8
ms.date: 06/08/2017
localization_priority: Normal
---


# ContentControlListEntry.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

 _expression_ An expression that returns an '[ContentControlListEntry](Word.ContentControlListEntry.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[ContentControlListEntry Object](Word.ContentControlListEntry.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]