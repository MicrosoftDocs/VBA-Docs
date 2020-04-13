---
title: Application.IsSandboxed property (Word)
keywords: vbawd10.chm158335468
f1_keywords:
- vbawd10.chm158335468
ms.prod: word
api_name:
- Word.Application.IsSandboxed
ms.assetid: 13fbfbda-b9e5-4f5d-46e2-2d8b157c77a1
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.IsSandboxed property (Word)

 **True** if the application window is a Protected View window. Read-only.


## Syntax

_expression_. `IsSandboxed`

 _expression_ An expression that returns a [Application](./Word.Application.md) object.


## Remarks

Use the **IsSandboxed** property to determine if a document is open within a Protected View window.


## Example

The following code example displays whether the specified document is open in a Protected View window.


```vb
Sub CheckIfSandboxed(doc As Document) 
 MsgBox doc.Application.IsSandboxed 
End Sub
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]