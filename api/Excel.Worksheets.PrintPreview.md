---
title: Worksheets.PrintPreview method (Excel)
keywords: vbaxl10.chm470082
f1_keywords:
- vbaxl10.chm470082
ms.prod: excel
api_name:
- Excel.Worksheets.PrintPreview
ms.assetid: cf0206e2-5016-2472-be89-c45e9c7a38f0
ms.date: 05/18/2019
localization_priority: Normal
---


# Worksheets.PrintPreview method (Excel)

Shows a preview of the object as it would look when printed.


## Syntax

_expression_.**PrintPreview** (_EnableChanges_)

_expression_ A variable that represents a **[Worksheets](Excel.Worksheets.md)** object.


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