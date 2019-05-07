---
title: LegendKey.Creator property (Word)
keywords: vbawd10.chm266207381
f1_keywords:
- vbawd10.chm266207381
ms.prod: word
api_name:
- Word.LegendKey.Creator
ms.assetid: ec1942b0-1ba3-cb55-1e0f-1bb8258f4810
ms.date: 06/08/2017
localization_priority: Normal
---


# LegendKey.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents a '[LegendKey](Word.LegendKey.md)' object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD". This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Word has the creator code MSWD. For more information about this property, consult the language reference Help included with Microsoft Office for Mac.


## See also


[LegendKey Object](Word.LegendKey.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]