---
title: PlotArea.Application property (Word)
keywords: vbawd10.chm53477524
f1_keywords:
- vbawd10.chm53477524
ms.prod: word
api_name:
- Word.PlotArea.Application
ms.assetid: 7c5a4198-6a1a-1039-7bf7-b88f4aa0a3e0
ms.date: 06/08/2017
localization_priority: Normal
---


# PlotArea.Application property (Word)

When used without an object qualifier, returns an  **[Application](Word.Application.md)** object that represents the Microsoft Word application. When used with an object qualifier, returns an **Application** object that represents the creator of the specified object (you can use this property with an Automation object to return the application of that object). Read-only.


## Syntax

 _expression_. `Application`

 _expression_ A variable that represents a '[PlotArea](Word.PlotArea.md)' object.


## Example

The following example displays a message about the application that created  `myObject`.


```vb
Set myObject = ActiveDocument 
If myObject.Application.Value = "Microsoft Word" Then 
 MsgBox "This is a Word Application object." 
Else 
 MsgBox "This is not a Word Application object." 
End If
```


## See also


[PlotArea Object](Word.PlotArea.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]