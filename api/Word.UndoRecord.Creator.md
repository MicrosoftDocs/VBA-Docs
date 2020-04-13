---
title: UndoRecord.Creator property (Word)
keywords: vbawd10.chm56099817
f1_keywords:
- vbawd10.chm56099817
ms.prod: word
api_name:
- Word.UndoRecord.Creator
ms.assetid: 8397c8ee-37d4-ec13-161a-4bd941226e2c
ms.date: 06/08/2017
localization_priority: Normal
---


# UndoRecord.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents an **[UndoRecord](Word.UndoRecord.md)** object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the **string** "MSWD". This property was primarily designed to be used on the Apple Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For more information about this property, see the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[UndoRecord Object](Word.UndoRecord.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]