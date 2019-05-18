---
title: Workbooks.Item property (Excel)
keywords: vbaxl10.chm203076
f1_keywords:
- vbaxl10.chm203076
ms.prod: excel
api_name:
- Excel.Workbooks.Item
ms.assetid: 2f01412d-8ba0-6911-81d3-e464a44354b5
ms.date: 05/18/2019
localization_priority: Normal
---


# Workbooks.Item property (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[Workbooks](Excel.Workbooks.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number of the object.|

## Example

This example sets the `wb` variable to the workbook for Myaddin.xla.

```vb
Set wb = Workbooks.Item("myaddin.xla")
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
