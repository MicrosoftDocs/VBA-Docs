---
title: ProtectedViewWindows.Creator property (Word)
keywords: vbawd10.chm82314217
f1_keywords:
- vbawd10.chm82314217
ms.prod: word
api_name:
- Word.ProtectedViewWindows.Creator
ms.assetid: 7de9abbc-e8b9-1f92-4f31-a8c8b1551106
ms.date: 06/08/2017
localization_priority: Normal
---


# ProtectedViewWindows.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ An expression that returns a **[ProtectedViewWindows](Word.ProtectedViewWindows.md)** object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the **string** "MSWD". This property was primarily designed to be used on the Apple Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, see the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[ProtectedViewWindows Object](Word.ProtectedViewWindows.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]