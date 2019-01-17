---
title: ModelChanges.Application property (Excel)
keywords: vbaxl10.chm959073
f1_keywords:
- vbaxl10.chm959073
ms.prod: excel
ms.assetid: 4f2d358a-ed68-1b9d-8eeb-e502a02d0c7f
ms.date: 06/08/2017
localization_priority: Normal
---


# ModelChanges.Application property (Excel)

Returns an  **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

_expression_. `Application`

_expression_ A variable that represents a 'ModelChanges' object.


## Example

This example displays a message about the application that created  `myObject`.


```vb
Set myObject = ActiveWorkbook 
If myObject.Application.Value = "Microsoft Excel" Then 
 MsgBox "This is an Excel Application object." 
Else 
 MsgBox "This is not an Excel Application object." 
End If
```


## Property value

 **APPLICATION**


## See also



[ModelChanges Object](Excel.modelchanges.md)

