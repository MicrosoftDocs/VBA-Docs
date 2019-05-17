---
title: Styles.Merge method (Excel)
keywords: vbaxl10.chm179076
f1_keywords:
- vbaxl10.chm179076
ms.prod: excel
api_name:
- Excel.Styles.Merge
ms.assetid: b2212f10-c16b-7108-8281-1c0375448f6d
ms.date: 05/16/2019
localization_priority: Normal
---


# Styles.Merge method (Excel)

Merges the styles from another workbook into the **Styles** collection.


## Syntax

_expression_.**Merge** (_Workbook_)

_expression_ A variable that represents a **[Styles](Excel.Styles.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Workbook_|Required| **Variant**|A **[Workbook](Excel.Workbook.md)** object that represents the workbook that contains styles to be merged.|

## Return value

Variant


## Example

This example merges the styles from the workbook Template.xls into the active workbook.

```vb
ActiveWorkbook.Styles.Merge Workbook:=Workbooks("TEMPLATE.XLS")
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]