---
title: Application.DoubleClick method (Excel)
keywords: vbaxl10.chm133128
f1_keywords:
- vbaxl10.chm133128
ms.prod: excel
api_name:
- Excel.Application.DoubleClick
ms.assetid: 17958601-3e24-a7bb-7d8c-0c45b955f449
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.DoubleClick method (Excel)

Equivalent to double-clicking the active cell.


## Syntax

_expression_.**DoubleClick**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Example

This example double-clicks the active cell on Sheet1.

```vb
Worksheets("Sheet1").Activate 
Application.DoubleClick
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
