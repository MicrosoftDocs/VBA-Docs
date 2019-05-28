---
title: Workbook.PrintPreview method (Excel)
keywords: vbaxl10.chm199128
f1_keywords:
- vbaxl10.chm199128
ms.prod: excel
api_name:
- Excel.Workbook.PrintPreview
ms.assetid: 044afc4c-74d6-3ea6-1811-2c7d9cdc5b1a
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.PrintPreview method (Excel)

Shows a preview of the object as it would look when printed.


## Syntax

_expression_.**PrintPreview** (_EnableChanges_)

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _EnableChanges_|Optional| **Variant**|Pass a **Boolean** value to specify if the user can change the margins and other page setup options available in print preview.|

## Example

This example displays Sheet1 in print preview.

```vb
Worksheets("Sheet1").PrintPreview
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]