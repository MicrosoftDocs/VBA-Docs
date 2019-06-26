---
title: StyleSheet.Path property (Word)
keywords: vbawd10.chm166658052
f1_keywords:
- vbawd10.chm166658052
ms.prod: word
api_name:
- Word.StyleSheet.Path
ms.assetid: 96a68487-b1b8-4c45-1869-b066874df9e5
ms.date: 06/08/2017
localization_priority: Normal
---


# StyleSheet.Path property (Word)

Returns the disk or Web path to the specified style sheet. Read-only  **String**.


## Syntax

_expression_.**Path**

_expression_ Required. A variable that represents a '[StyleSheet](Word.StyleSheet.md)' object.


## Remarks

The path doesn't include a trailing character—for example, "C:\MSOffice" or "https://MyServer". Use the  **[PathSeparator](Word.Application.PathSeparator.md)** property to add the character that separates folders and drive letters, and use the **[Name](Word.StyleSheet.Name.md)** property to return the file name without the path.


> [!NOTE] 
> You can use the  **PathSeparator** property to build web addresses even though they contain forward slashes (/) and the **PathSeparator** property defaults to a backslash (\).


## See also


[StyleSheet Object](Word.StyleSheet.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]