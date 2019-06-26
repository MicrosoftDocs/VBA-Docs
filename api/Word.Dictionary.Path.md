---
title: Dictionary.Path property (Word)
keywords: vbawd10.chm162332673
f1_keywords:
- vbawd10.chm162332673
ms.prod: word
api_name:
- Word.Dictionary.Path
ms.assetid: 1fd2d6ac-e112-9d13-0e41-2584e6841b73
ms.date: 06/08/2017
localization_priority: Normal
---


# Dictionary.Path property (Word)

Returns the path to the specified dictionary. Read-only  **String**.


## Syntax

_expression_.**Path**

_expression_ Required. A variable that represents a '[Dictionary](Word.Dictionary.md)' object.


## Remarks

The path doesn't include a trailing character — for example, "C:\MSOffice" or "https://MyServer". Use the  **PathSeparator** property to add the character that separates folders and drive letters. Use the **Name** property to return the file name without the path and use the **FullName** property to return the file name and the path together.


> [!NOTE] 
> You can use the  **PathSeparator** property to build web addresses even though they contain forward slashes (/) and the **PathSeparator** property defaults to a backslash (\).


## See also


[Dictionary Object](Word.Dictionary.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]