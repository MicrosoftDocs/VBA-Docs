---
title: Application.ExtendList property (Excel)
keywords: vbaxl10.chm133243
f1_keywords:
- vbaxl10.chm133243
ms.prod: excel
api_name:
- Excel.Application.ExtendList
ms.assetid: b368047b-9d30-5a6f-a7db-748e3e91a3c0
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.ExtendList property (Excel)

**True** if Microsoft Excel automatically extends formatting and formulas to new data that is added to a list. Read/write **Boolean**.


## Syntax

_expression_.**ExtendList**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Remarks

To be extended, formats and formulas must appear in at least three of the five list rows or columns preceding the new row or column, and you must add the data to the bottom or to the right side of the list.


## Example

This example sets Excel to not apply formatting and formulas to data subsequently added to an existing list.

```vb
Application.ExtendList = False
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]