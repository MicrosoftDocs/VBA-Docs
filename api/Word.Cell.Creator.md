---
title: Cell.Creator property (Word)
keywords: vbawd10.chm156107753
f1_keywords:
- vbawd10.chm156107753
ms.prod: word
api_name:
- Word.Cell.Creator
ms.assetid: 9a50df51-61ab-01d1-30fe-6c5f6622ce4c
ms.date: 06/08/2017
localization_priority: Normal
---


# Cell.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a '[Cell](Word.Cell.md)' object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[Cell Object](Word.Cell.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]