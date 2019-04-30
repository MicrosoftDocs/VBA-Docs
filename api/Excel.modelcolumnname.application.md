---
title: ModelColumnName.Application property (Excel)
keywords: vbaxl10.chm961073
f1_keywords:
- vbaxl10.chm961073
ms.prod: excel
ms.assetid: a15b21c5-0d29-8e5c-2d85-0d8d5810fba1
ms.date: 05/01/2019
localization_priority: Normal
---


# ModelColumnName.Application property (Excel)

Returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[ModelColumnName](Excel.modelcolumnname.md)** object.


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