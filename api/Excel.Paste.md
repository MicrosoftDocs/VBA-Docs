---
title: Paste method (Excel Graph)
keywords: vbagr10.chm5207755
f1_keywords:
- vbagr10.chm5207755
ms.prod: excel
api_name:
- Excel.Paste
ms.assetid: 4cb4fa45-b319-f3a8-e477-80b96060905b
ms.date: 04/09/2019
localization_priority: Normal
---


# Paste method (Excel Graph)

Pastes the contents of the Clipboard into the specified range on the datasheet.

## Syntax

_expression_.**Paste** (_Link_)

_expression_ Required. An expression that returns a **[Range](excel.range-graph-object.md)** object.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Link_ | Optional |**Variant**| **True** to establish a link to the source of the pasted data. The default value is **False**.|

## Example

This example pastes the contents of the Clipboard into cell A1 on the datasheet.

```vb
myChart.Application.DataSheet.Range("A1").Paste
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]