---
title: Export method (Excel Graph)
keywords: vbagr10.chm66950
f1_keywords:
- vbagr10.chm66950
ms.prod: excel
api_name:
- Excel.Export
ms.assetid: c5929fa7-ac8a-0cb4-5c8d-01e5cfa23d1e
ms.date: 04/09/2019
localization_priority: Normal
---


# Export method (Excel Graph)

Exports the chart in a graphic format. Returns a value of type **Boolean**.

## Syntax

_expression_.**Export** (_FileName_, _FilterName_, _Interactive_)

_expression_ Required. An expression that returns a **[Chart](Excel.Chart-graph-object.md)** object.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_FileName_ |Required |**String**|The name of the exported file.|
|_FilterName_ |Optional |**Variant**|The language-independent name of the graphic filter as it appears in the registry.|
|_Interactive_ |Optional |**Variant**|**True** to display the dialog box that contains the filter-specific options. If this argument is **False**, Graph uses the default values for the filter. The default value is **False**.|

## Example

This example exports the chart as a GIF file.

```vb
myChart.Export _ 
 FileName:="current_sales.gif", FilterName:="GIF"
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]