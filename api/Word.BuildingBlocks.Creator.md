---
title: BuildingBlocks.Creator property (Word)
ms.prod: word
api_name:
- Word.BuildingBlocks.Creator
ms.assetid: 42d378dc-d442-e8e2-382c-ea82b71ffcf8
ms.date: 06/08/2017
localization_priority: Normal
---


# BuildingBlocks.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

 _expression_ An expression that returns an '[BuildingBlocks](Word.BuildingBlocks.md)' object.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[BuildingBlocks Collection](Word.BuildingBlocks.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]