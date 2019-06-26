---
title: Subdocument.Path property (Word)
keywords: vbawd10.chm159973380
f1_keywords:
- vbawd10.chm159973380
ms.prod: word
api_name:
- Word.Subdocument.Path
ms.assetid: d27bc7ce-5346-b9a7-cd29-b42b0e8c98eb
ms.date: 06/08/2017
localization_priority: Normal
---


# Subdocument.Path property (Word)

Returns the disk or Web path to the specified subdocument. Read-only  **String**.


## Syntax

_expression_.**Path**

_expression_ Required. A variable that represents a '[Subdocument](Word.Subdocument.md)' object.


## Remarks

The path doesn't include a trailing character — for example, "C:\MSOffice" or "https://MyServer". Use the  **[PathSeparator](Word.Application.PathSeparator.md)** property to add the character that separates folders and drive letters. Use the **[Name](Word.Subdocument.Name.md)** property to return the file name without the path.


> [!NOTE] 
> You can use the  **PathSeparator** property to build web addresses even though they contain forward slashes (/) and the **PathSeparator** property defaults to a backslash (\).


## See also


[Subdocument Object](Word.Subdocument.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]