---
title: Gridlines.Application property (Word)
keywords: vbawd10.chm11468948
f1_keywords:
- vbawd10.chm11468948
ms.prod: word
api_name:
- Word.Gridlines.Application
ms.assetid: 49950209-0edd-e2b2-a2a7-35a6ce9dc0cb
ms.date: 06/08/2017
localization_priority: Normal
---


# Gridlines.Application property (Word)

When used without an object qualifier, returns an  **[Application](Word.Application.md)** object that represents the Microsoft Word application. When used with an object qualifier, returns an **Application** object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a '[GridLines](Word.GridLines.md)' object.


## Example

The following example displays a message about the application that created _myObject_.


```vb
Set myObject = ActiveDocument 
If myObject.Application.Value = "Microsoft Word" Then 
 MsgBox "This is a Word Application object." 
Else 
 MsgBox "This is not a Word Application object." 
End If
```


## See also


[GridLines Object](Word.GridLines.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]