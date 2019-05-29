---
title: Workbook.RemoveUser method (Excel)
keywords: vbaxl10.chm199138
f1_keywords:
- vbaxl10.chm199138
ms.prod: excel
api_name:
- Excel.Workbook.RemoveUser
ms.assetid: f0a978a0-7bcf-3af4-a01a-831c6c854989
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.RemoveUser method (Excel)

Disconnects the specified user from the shared workbook.


## Syntax

_expression_.**RemoveUser** (_Index_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The user index.|

## Example

This example disconnects user two from the shared workbook.

```vb
Workbooks(2).RemoveUser 2
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]