---
title: LegendEntries.Creator property (Word)
keywords: vbawd10.chm6815893
f1_keywords:
- vbawd10.chm6815893
ms.prod: word
api_name:
- Word.LegendEntries.Creator
ms.assetid: 4153e082-a1df-a118-3db6-4603e991bf9c
ms.date: 06/08/2017
localization_priority: Normal
---


# LegendEntries.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a '[LegendEntries](Word.LegendEntries.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD". This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


[LegendEntries Object](Word.LegendEntries.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]