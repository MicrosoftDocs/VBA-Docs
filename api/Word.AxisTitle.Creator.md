---
title: AxisTitle.Creator property (Word)
keywords: vbawd10.chm98238613
f1_keywords:
- vbawd10.chm98238613
ms.prod: word
api_name:
- Word.AxisTitle.Creator
ms.assetid: 681851d1-b045-85a1-e4bc-9fface4d4b00
ms.date: 06/08/2017
localization_priority: Normal
---


# AxisTitle.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents an '[AxisTitle](Word.AxisTitle.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


[AxisTitle Object](Word.AxisTitle.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]