---
title: BuildingBlockType.Creator property (Word)
keywords: vbawd10.chm167379945
f1_keywords:
- vbawd10.chm167379945
ms.prod: word
api_name:
- Word.BuildingBlockType.Creator
ms.assetid: 6c242dbd-94ea-2ac1-5dc9-3118b5453d01
ms.date: 06/08/2017
localization_priority: Normal
---


# BuildingBlockType.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the add-in was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

 _expression_ An expression that returns an '[BuildingBlockType](Word.BuildingBlockType.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[BuildingBlockType Object](Word.BuildingBlockType.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]