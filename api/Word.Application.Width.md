---
title: Application.Width property (Word)
keywords: vbawd10.chm158335065
f1_keywords:
- vbawd10.chm158335065
ms.prod: word
api_name:
- Word.Application.Width
ms.assetid: ac9b369e-6661-ef67-6674-85ab02ef4621
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Width property (Word)

Returns or sets the width of the application window, in points. Read/write  **Long**.


## Syntax

_expression_.**Width**

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 


## Example

This example sets the width and height of the Microsoft Word application window.


```vb
With Application 
 .WindowState = wdWindowStateNormal 
 .Width = 500 
 .Height = 400 
End With
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]