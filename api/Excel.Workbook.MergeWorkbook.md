---
title: Workbook.MergeWorkbook method (Excel)
keywords: vbaxl10.chm199111
f1_keywords:
- vbaxl10.chm199111
ms.prod: excel
api_name:
- Excel.Workbook.MergeWorkbook
ms.assetid: 393790c6-3c19-7149-a999-b8712e7a6855
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.MergeWorkbook method (Excel)

Merges changes from one workbook into an open workbook.


## Syntax

_expression_.**MergeWorkbook** (_FileName_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required| **Variant**|The file name of the workbook that contains the changes to be merged into the open workbook.|

## Example

This example merges changes from Book1.xls into the active workbook.

```vb
ActiveWorkbook.MergeWorkbook "Book1.xls"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]