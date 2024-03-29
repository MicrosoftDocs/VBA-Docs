---
title: ModelTable.Application property (Excel)
keywords: vbaxl10.chm933073
f1_keywords:
- vbaxl10.chm933073
ms.assetid: c2138114-e623-f141-090d-22644f8d2477
ms.date: 05/01/2019
ms.localizationpriority: medium
---


# ModelTable.Application property (Excel)

Returns an **[Application](Excel.Application(object).md)** object that represents the Microsoft Excel application. Read-only.


## Syntax

_expression_.**Application**

_expression_ A variable that represents a **[ModelTable](Excel.modeltable.md)** object.


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