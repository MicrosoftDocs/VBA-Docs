---
title: Copy method (Excel Graph)
keywords: vbagr10.chm66087
f1_keywords:
- vbagr10.chm66087
ms.prod: excel
api_name:
- Excel.Copy
ms.assetid: 2207804d-0003-5c75-afa8-a718efba0c2c
ms.date: 04/09/2019
localization_priority: Normal
---


# Copy method (Excel Graph)

The **Copy** method as it applies to the **ChartArea** and **Range** objects.

## ChartArea object

Copies a picture of the point or series to the Clipboard.

### Syntax

_expression_.**Copy**

_expression_ Required. An expression that returns a **[ChartArea](excel.chartarea-graph-object.md)** object.




## Range object

Copies the range to the specified range or to the Clipboard.

### Syntax

_expression_.**Copy** (_Destination_)

_expression_ Required. An expression that returns a **[Range](excel.range-graph-object.md)** object. 

### Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Destination_| Optional |**Variant**|Specifies the new range to which the specified range will be copied. If this argument is omitted, Graph copies the range to the Clipboard.|

## Example

This example copies the formulas in cells A1:D4 on the datasheet into cells E5:H8.

```vb
Set mySheet = myChart.Application.DataSheet 
mySheet.Range("A1:D4").Copy _ 
 Destination:= mySheet.Range("E5")
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]