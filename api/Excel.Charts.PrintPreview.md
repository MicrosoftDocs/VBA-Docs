---
title: Charts.PrintPreview method (Excel)
keywords: vbaxl10.chm217080
f1_keywords:
- vbaxl10.chm217080
ms.prod: excel
api_name:
- Excel.Charts.PrintPreview
ms.assetid: 53d54413-6c35-d2a3-ba4a-1acc3bbdea28
ms.date: 04/20/2019
localization_priority: Normal
---


# Charts.PrintPreview method (Excel)

Shows a preview of the object as it would look when printed.


## Syntax

_expression_.**PrintPreview** (_EnableChanges_)

_expression_ A variable that represents a **[Charts](Excel.Charts.md)** object.


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