---
title: ChartColorFormat.Application property (Word)
keywords: vbawd10.chm12058772
f1_keywords:
- vbawd10.chm12058772
ms.prod: word
api_name:
- Word.ChartColorFormat.Application
ms.assetid: ca8d6709-c65b-fca9-1061-b23c6107eba2
ms.date: 06/08/2017
localization_priority: Normal
---


# ChartColorFormat.Application property (Word)

When used without an object qualifier, returns an  **[Application](Word.Application.md)** object that represents the Microsoft Word application. When used with an object qualifier, returns an **Application** object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a '[ChartColorFormat](Word.ChartColorFormat.md)' object.


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


[ChartColorFormat Object](Word.ChartColorFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]