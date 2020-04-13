---
title: Template.Path property (Word)
keywords: vbawd10.chm157941761
f1_keywords:
- vbawd10.chm157941761
ms.prod: word
api_name:
- Word.Template.Path
ms.assetid: 9b84e053-b806-d43d-2c3c-b8ce56cf7d15
ms.date: 06/08/2017
localization_priority: Normal
---


# Template.Path property (Word)

Returns the path to the specified document template. Read-only  **String**.


## Syntax

_expression_.**Path**

_expression_ Required. A variable that represents a '[Template](Word.Template.md)' object.


## Remarks

The path doesn't include a trailing character — for example, "C:\MSOffice" or "https://MyServer". Use the **[PathSeparator](Word.Application.PathSeparator.md)** property to add the character that separates folders and drive letters. Use the **[Name](Word.Template.Name.md)** property to return the file name without the path and use the **[FullName](Word.Template.FullName.md)** property to return the file name and the path together.


## See also


[Template Object](Word.Template.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]