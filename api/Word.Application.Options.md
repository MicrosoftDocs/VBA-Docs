---
title: Application.Options property (Word)
keywords: vbawd10.chm158335069
f1_keywords:
- vbawd10.chm158335069
ms.prod: word
api_name:
- Word.Application.Options
ms.assetid: 87bf2092-8707-d375-d4d6-f7420be1fe7d
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Options property (Word)

Returns an  **[Options](Word.Options.md)** object that represents application settings in Microsoft Word.


## Syntax

 _expression_. `Options`

 _expression_ A variable that represents an '[Application](Word.Application.md)' object.


## Example

This example disables fast saves and then saves the active document.


```vb
Options.AllowFastSave = False 
ActiveDocument.Save
```

This example prints Sales.doc with comments and field results.




```vb
With Options 
 .PrintFieldCodes = False 
 .PrintComments = True 
End With 
Documents("Sales.doc").PrintOut
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]