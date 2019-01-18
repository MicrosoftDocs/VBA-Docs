---
title: AddIns.Creator property (Word)
keywords: vbawd10.chm159319017
f1_keywords:
- vbawd10.chm159319017
ms.prod: word
api_name:
- Word.AddIns.Creator
ms.assetid: 9789df8f-fc50-32b3-50a2-39a540eeacb1
ms.date: 06/08/2017
localization_priority: Normal
---


# AddIns.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

 _expression_. `Creator`

 _expression_ An expression that returns an '[AddIns](Word.addins.md)' collection.


## Remarks

If the object was created in Microsoft Word, the  **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


 **Note**  This value can also be represented by the constant  **wdCreatorCode**.


## See also


[AddIns Collection Object](Word.addins.md)

