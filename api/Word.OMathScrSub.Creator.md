---
title: OMathScrSub.Creator property (Word)
keywords: vbawd10.chm219283557
f1_keywords:
- vbawd10.chm219283557
ms.prod: word
api_name:
- Word.OMathScrSub.Creator
ms.assetid: 9919a737-e295-590e-021f-911a7146fa73
ms.date: 06/08/2017
localization_priority: Normal
---


# OMathScrSub.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

 _expression_ An expression that returns an '[OMathScrSub](Word.OMathScrSub.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[OMathScrSub Object](Word.OMathScrSub.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]