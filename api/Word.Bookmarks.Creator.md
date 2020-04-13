---
title: Bookmarks.Creator property (Word)
keywords: vbawd10.chm157746153
f1_keywords:
- vbawd10.chm157746153
ms.prod: word
api_name:
- Word.Bookmarks.Creator
ms.assetid: 897e3571-2147-c92c-e07a-0bd0d75b107b
ms.date: 06/08/2017
localization_priority: Normal
---


# Bookmarks.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents an '[Bookmarks](Word.bookmarks.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[Bookmarks Collection Object](Word.bookmarks.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]