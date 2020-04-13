---
title: OMathFunction.Creator property (Word)
keywords: vbawd10.chm22151269
f1_keywords:
- vbawd10.chm22151269
ms.prod: word
api_name:
- Word.OMathFunction.Creator
ms.assetid: bdf1a3d3-7aad-25ad-7407-b051d7b765c6
ms.date: 06/08/2017
localization_priority: Normal
---


# OMathFunction.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

 _expression_ An expression that returns an '[OMathFunction](Word.OMathFunction.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[OMathFunction Object](Word.OMathFunction.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]