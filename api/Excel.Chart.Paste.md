---
title: Chart.Paste method (Excel)
keywords: vbaxl10.chm149129
f1_keywords:
- vbaxl10.chm149129
ms.prod: excel
api_name:
- Excel.Chart.Paste
ms.assetid: e34d3d30-39f8-dbd4-1a39-d3ef9f84e0f4
ms.date: 04/16/2019
localization_priority: Normal
---


# Chart.Paste method (Excel)

Pastes chart data from the Clipboard into the specified chart.


## Syntax

_expression_.**Paste** (_Type_)

_expression_ A variable that represents a **[Chart](Excel.Chart(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Optional| **Variant**|Specifies the chart information to paste if a chart is on the Clipboard. Can be one of the following **[XlPasteType](Excel.XlPasteType.md)** constants: **xlPasteFormats**, **xlPasteFormulas**, or **xlPasteAll**. The default value is **xlPasteAll**. If there's data other than a chart on the Clipboard, this argument cannot be used.|

## Example

This example pastes data from the range B1:B5 on Sheet1 into Chart1.

```vb
Worksheets("Sheet1").Range("B1:B5").Copy 
Charts("Chart1").Paste
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
