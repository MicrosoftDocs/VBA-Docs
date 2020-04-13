---
title: ListTemplates.Creator property (Word)
keywords: vbawd10.chm160433129
f1_keywords:
- vbawd10.chm160433129
ms.prod: word
api_name:
- Word.ListTemplates.Creator
ms.assetid: 8af8b8e3-fce0-3770-01de-12bea22b6792
ms.date: 06/08/2017
localization_priority: Normal
---


# ListTemplates.Creator property (Word)

Returns a 32-bit integer that indicates the application in which the specified object was created. Read-only  **Long**.


## Syntax

_expression_.**Creator**

_expression_ Required. A variable that represents a '[ListTemplates](Word.listtemplates.md)' collection.


## Remarks

If the object was created in Microsoft Word, the **Creator** property returns the hexadecimal number 4D535744, which represents the string "MSWD." This property was primarily designed to be used on the Macintosh, where each application has a four-character creator code. For example, Microsoft Word has the creator code MSWD. For additional information about this property, consult the language reference Help included with Microsoft Office Macintosh Edition.


## See also


[ListTemplates Collection Object](Word.listtemplates.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]