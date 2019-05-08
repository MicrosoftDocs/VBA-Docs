---
title: TickLabels.Creator property (Word)
keywords: vbawd10.chm167051413
f1_keywords:
- vbawd10.chm167051413
ms.prod: word
api_name:
- Word.TickLabels.Creator
ms.assetid: 854570ae-1e01-7b32-8c2d-8643c8912b82
ms.date: 06/08/2017
localization_priority: Normal
---


# TickLabels.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a '[TickLabels](Word.TickLabels.md)' object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD". This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


[TickLabels Object](Word.TickLabels.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]