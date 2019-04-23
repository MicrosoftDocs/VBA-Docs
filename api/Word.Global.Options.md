---
title: Global.Options property (Word)
keywords: vbawd10.chm163119197
f1_keywords:
- vbawd10.chm163119197
ms.prod: word
api_name:
- Word.Global.Options
ms.assetid: 1d73dd2d-2fdd-7f12-ce6d-c6b7542d284c
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.Options property (Word)

Returns an  **Options** object that represents application settings in Microsoft Word.


## Syntax

_expression_. `Options`

_expression_ Required. A variable that represents a '[Global](Word.Global.md)' object.


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


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]