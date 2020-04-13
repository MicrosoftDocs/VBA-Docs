---
title: CoAuthUpdates.Creator property (Word)
keywords: vbawd10.chm217842665
f1_keywords:
- vbawd10.chm217842665
ms.prod: word
api_name:
- Word.CoAuthUpdates.Creator
ms.assetid: abd8a680-050b-7866-c198-c2e258281bc9
ms.date: 06/08/2017
localization_priority: Normal
---


# CoAuthUpdates.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

 _expression_ An expression that returns a 'CoAuthUpdates' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the **string** "MSWD". This property was primarily designed to be used on the Apple Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For more information about this property, see the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[CoAuthUpdates Object](overview/Word.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]