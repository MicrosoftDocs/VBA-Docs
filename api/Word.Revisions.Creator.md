---
title: Revisions.Creator property (Word)
keywords: vbawd10.chm159384553
f1_keywords:
- vbawd10.chm159384553
ms.prod: word
api_name:
- Word.Revisions.Creator
ms.assetid: c8db3880-70c4-7d3f-5705-828e061f2c52
ms.date: 06/08/2017
localization_priority: Normal
---


# Revisions.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

 _expression_ An expression that returns a 'Revisions' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the **string** "MSWD". This property was primarily designed to be used on the Apple Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For more information about this property, see the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[Revisions Collection Object](Word.revisions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]