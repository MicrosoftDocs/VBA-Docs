---
title: OMathBorderBox.Creator property (Word)
keywords: vbawd10.chm116260965
f1_keywords:
- vbawd10.chm116260965
ms.prod: word
api_name:
- Word.OMathBorderBox.Creator
ms.assetid: 816122c9-fa58-b3a4-ff25-74e0ad6b7015
ms.date: 06/08/2017
localization_priority: Normal
---


# OMathBorderBox.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

 _expression_ An expression that returns an '[OMathBorderBox](Word.OMathBorderBox.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[OMathBorderBox Object](Word.OMathBorderBox.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]