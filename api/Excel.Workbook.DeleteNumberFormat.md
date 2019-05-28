---
title: Workbook.DeleteNumberFormat method (Excel)
keywords: vbaxl10.chm199096
f1_keywords:
- vbaxl10.chm199096
ms.prod: excel
api_name:
- Excel.Workbook.DeleteNumberFormat
ms.assetid: d56c2e4c-5de2-fecf-6a1f-a9fdc79943cb
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.DeleteNumberFormat method (Excel)

Deletes a custom number format from the workbook.


## Syntax

_expression_.**DeleteNumberFormat** (_NumberFormat_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _NumberFormat_|Required| **String**|Names the number format to be deleted.|

## Example

This example deletes the number format "000-00-0000" from the active workbook.

```vb
ActiveWorkbook.DeleteNumberFormat("000-00-0000")
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]