---
title: AutoCaptions.Creator property (Word)
keywords: vbawd10.chm158991337
f1_keywords:
- vbawd10.chm158991337
ms.prod: word
api_name:
- Word.AutoCaptions.Creator
ms.assetid: 998c1603-210a-bc79-47d5-f3138ea09d8d
ms.date: 06/08/2017
localization_priority: Normal
---


# AutoCaptions.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ A variable that represents an '[AutoCaptions](Word.autocaptions.md)' object.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


> [!NOTE] 
> This value can also be represented by the constant **wdCreatorCode**.


## See also


[AutoCaptions Collection Object](Word.autocaptions.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]