---
title: Application.MouseAvailable property (Excel)
keywords: vbaxl10.chm133167
f1_keywords:
- vbaxl10.chm133167
ms.prod: excel
api_name:
- Excel.Application.MouseAvailable
ms.assetid: b22f9d44-6a84-6716-d663-450f08c5557d
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.MouseAvailable property (Excel)

**True** if a mouse is available. Read-only **Boolean**.


## Syntax

_expression_.**MouseAvailable**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example displays a message if a mouse isn't available.

```vb
If Application.MouseAvailable = False Then 
 MsgBox "Your system does not have a mouse" 
End If
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]