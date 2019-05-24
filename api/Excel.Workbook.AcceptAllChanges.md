---
title: Workbook.AcceptAllChanges method (Excel)
keywords: vbaxl10.chm199177
f1_keywords:
- vbaxl10.chm199177
ms.prod: excel
api_name:
- Excel.Workbook.AcceptAllChanges
ms.assetid: 8d8572a9-1231-c8ef-0707-72b8b5109066
ms.date: 05/25/2019
localization_priority: Normal
---


# Workbook.AcceptAllChanges method (Excel)

Accepts all changes in the specified shared workbook.


## Syntax

_expression_.**AcceptAllChanges** (_When_, _Who_, _Where_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _When_|Optional| **Variant**|Specifies when all the changes are accepted.|
| _Who_|Optional| **Variant**|Specifies by whom all the changes are accepted.|
| _Where_|Optional| **Variant**|Specifies where all the changes are accepted.|

## Example

This example accepts all changes in the active workbook.

```vb
ActiveWorkbook.AcceptAllChanges
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]