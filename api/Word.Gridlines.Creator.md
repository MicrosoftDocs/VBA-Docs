---
title: Gridlines.Creator property (Word)
keywords: vbawd10.chm11468949
f1_keywords:
- vbawd10.chm11468949
ms.prod: word
api_name:
- Word.Gridlines.Creator
ms.assetid: e3548aa9-1e88-b5ee-eb4a-417504d7bde5
ms.date: 06/08/2017
localization_priority: Normal
---


# Gridlines.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a '[GridLines](Word.GridLines.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD". This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


[GridLines Object](Word.GridLines.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]