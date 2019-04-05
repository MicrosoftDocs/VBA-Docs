---
title: Application.OrganizationName property (Excel)
keywords: vbaxl10.chm133188
f1_keywords:
- vbaxl10.chm133188
ms.prod: excel
api_name:
- Excel.Application.OrganizationName
ms.assetid: 4255a006-52df-66f6-2948-a9522e3adfef
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.OrganizationName property (Excel)

Returns the registered organization name. Read-only **String**.


## Syntax

_expression_.**OrganizationName**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example displays the registered organization name.

```vb
MsgBox "The registered organization is " & _ 
 Application.OrganizationName
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]