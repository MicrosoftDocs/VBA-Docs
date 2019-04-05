---
title: Application.ProductCode property (Excel)
keywords: vbaxl10.chm133248
f1_keywords:
- vbaxl10.chm133248
ms.prod: excel
api_name:
- Excel.Application.ProductCode
ms.assetid: 5fd20091-4c74-f39c-9842-a5a032048edd
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.ProductCode property (Excel)

Returns the globally unique identifier (GUID) for Microsoft Excel. Read-only **String**.


## Syntax

_expression_.**ProductCode**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example displays the GUID for Excel.

```vb
MsgBox Application.ProductCode
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]