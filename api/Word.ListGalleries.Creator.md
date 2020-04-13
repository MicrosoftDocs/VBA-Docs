---
title: ListGalleries.Creator property (Word)
keywords: vbawd10.chm160760809
f1_keywords:
- vbawd10.chm160760809
ms.prod: word
api_name:
- Word.ListGalleries.Creator
ms.assetid: 2c24a4a7-b109-0b50-483a-b118b76ed731
ms.date: 06/08/2017
localization_priority: Normal
---


# ListGalleries.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

 _expression_ An expression that returns a 'ListGalleries' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the **string** "MSWD". This property was primarily designed to be used on the Apple Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For more information about this property, see the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[ListGalleries Collection Object](Word.listgalleries.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]