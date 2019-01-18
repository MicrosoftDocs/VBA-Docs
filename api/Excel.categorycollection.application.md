---
title: CategoryCollection.Application property (Excel)
keywords: vbaxl10.chm947073
f1_keywords:
- vbaxl10.chm947073
ms.prod: excel
ms.assetid: cfae4e60-9cda-c43b-e1d5-78ba110dd21c
ms.date: 06/08/2017
localization_priority: Normal
---


# CategoryCollection.Application property (Excel)

Returns an  **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

_expression_. `Application`

_expression_ A variable that represents a [CategoryCollection object (Excel)](Excel.categorycollection.md) object.


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



[CategoryCollection Object](Excel.categorycollection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]