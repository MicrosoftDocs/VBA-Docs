---
title: BuildingBlockEntries.Creator property (Word)
keywords: vbawd10.chm36242409
f1_keywords:
- vbawd10.chm36242409
ms.prod: word
api_name:
- Word.BuildingBlockEntries.Creator
ms.assetid: 41cc6b8f-18f1-695e-f811-fc597e6dcb51
ms.date: 06/08/2017
localization_priority: Normal
---


# BuildingBlockEntries.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

 _expression_ An expression that returns an '[BuildingBlockEntries](Word.BuildingBlockEntries.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[BuildingBlockEntries Collection](Word.BuildingBlockEntries.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]