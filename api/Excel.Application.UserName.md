---
title: Application.UserName property (Excel)
keywords: vbaxl10.chm133225
f1_keywords:
- vbaxl10.chm133225
ms.prod: excel
api_name:
- Excel.Application.UserName
ms.assetid: 6cb2636c-ef3c-82fb-583d-8390cc815631
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.UserName property (Excel)

Returns or sets the name of the current user. Read/write **String**.


## Syntax

_expression_.**UserName**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example displays the name of the current user.

```vb
MsgBox "Current user is " & Application.UserName
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
