---
title: CoAuthors.Creator property (Word)
keywords: vbawd10.chm179962857
f1_keywords:
- vbawd10.chm179962857
ms.prod: word
api_name:
- Word.CoAuthors.Creator
ms.assetid: a94deeeb-992f-ec40-9080-cb4aa6a6e1d5
ms.date: 06/08/2017
localization_priority: Normal
---


# CoAuthors.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

 _expression_ An expression that returns a [CoAuthors](./Word.CoAuthors.md) object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the **string** "MSWD". This property was primarily designed to be used on the Apple Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For more information about this property, see the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[CoAuthors Object](Word.CoAuthors.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]