---
title: OMathScrPre.Creator property (Word)
keywords: vbawd10.chm202244197
f1_keywords:
- vbawd10.chm202244197
ms.prod: word
api_name:
- Word.OMathScrPre.Creator
ms.assetid: e2c92e28-baed-17be-1631-f1aa2e357100
ms.date: 06/08/2017
localization_priority: Normal
---


# OMathScrPre.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

 _expression_ An expression that returns an '[OMathScrPre](Word.OMathScrPre.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[OMathScrPre Object](Word.OMathScrPre.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]