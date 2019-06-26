---
title: OMathRecognizedFunctions.Creator property (Word)
keywords: vbawd10.chm143065189
f1_keywords:
- vbawd10.chm143065189
ms.prod: word
api_name:
- Word.OMathRecognizedFunctions.Creator
ms.assetid: c1f81f2b-7a51-e6f7-ad62-e11088bd79ad
ms.date: 06/08/2017
localization_priority: Normal
---


# OMathRecognizedFunctions.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

 _expression_ An expression that returns an '[OMathRecognizedFunctions](Word.OMathRecognizedFunctions.md)' object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[OMathRecognizedFunctions Collection](Word.OMathRecognizedFunctions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]