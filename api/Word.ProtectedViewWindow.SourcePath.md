---
title: ProtectedViewWindow.SourcePath property (Word)
keywords: vbawd10.chm231735307
f1_keywords:
- vbawd10.chm231735307
ms.prod: word
api_name:
- Word.ProtectedViewWindow.SourcePath
ms.assetid: 05b4e601-894a-de8f-1119-565183b244b7
ms.date: 06/08/2017
localization_priority: Normal
---


# ProtectedViewWindow.SourcePath property (Word)

Returns the path of the source file for the specified Protected View window. Read-only  **String**.


## Syntax

_expression_.**SourcePath**

 _expression_ An expression that returns a [ProtectedViewWindow](./Word.ProtectedViewWindow.md) object.


## Remarks

The path does not include a trailing character (for example, "C:\MSOffice"). Use the [PathSeparator](Word.Application.PathSeparator.md) property to add the character that separates folders and drive letters. Use the [SourceName](Word.LinkFormat.SourceName.md) property to return the file name without the path.


## Example

The following code example returns the path and name of the document associated with the specified Protected View window.


```vb
MsgBox ActiveProtectedViewWindow.SourcePath & Application.PathSeparator _ 
 & ActiveProtectedViewWindow.SourceName 

```


## See also


[ProtectedViewWindow Object](Word.ProtectedViewWindow.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]