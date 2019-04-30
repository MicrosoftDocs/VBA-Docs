---
title: ModelColumnNames.Application property (Excel)
keywords: vbaxl10.chm963073
f1_keywords:
- vbaxl10.chm963073
ms.prod: excel
ms.assetid: 09a0a219-b4eb-4ead-f058-5b9a04e98dc9
ms.date: 05/01/2019
localization_priority: Normal
---


# ModelColumnNames.Application property (Excel)

Returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[ModelColumnNames](Excel.modelcolumnnames.md)** object.


## Property value

**APPLICATION**


## Example

This example displays a message about the application that created _myObject_.

```vb
Set myObject = ActiveWorkbook 
If myObject.Application.Value = "Microsoft Excel" Then 
 MsgBox "This is an Excel Application object." 
Else 
 MsgBox "This is not an Excel Application object." 
End If
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]