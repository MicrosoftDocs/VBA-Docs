---
title: Worksheet.PrintPreview method (Excel)
keywords: vbaxl10.chm174088
f1_keywords:
- vbaxl10.chm174088
ms.prod: excel
api_name:
- Excel.Worksheet.PrintPreview
ms.assetid: e7065877-2ec9-01ba-4672-4b5a0a8459d2
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.PrintPreview method (Excel)

Shows a preview of the object as it would look when printed.


## Syntax

_expression_.**PrintPreview** (_EnableChanges_)

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _EnableChanges_|Optional| **Variant**|Passes a **Boolean** value to specify if the user can change the margins and other page setup options available in print preview.|

## Example

This example displays Sheet1 in print preview.

```vb
Worksheets("Sheet1").PrintPreview
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]