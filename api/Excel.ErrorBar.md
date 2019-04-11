---
title: ErrorBar method (Excel Graph)
keywords: vbagr10.chm65688
f1_keywords:
- vbagr10.chm65688
ms.prod: excel
api_name:
- Excel.ErrorBar
ms.assetid: c2ada146-1549-aa88-2a39-bf1cccf1008b
ms.date: 04/09/2019
localization_priority: Normal
---


# ErrorBar method (Excel Graph)

Applies error bars to the specified series. **Variant**.

## Syntax

_expression_.**ErrorBar** (_Direction_, _Include_, _Type_, _Amount_, _MinusValues_)

_expression_ Required. An expression that returns one of the objects in the **Applies To** list.

## Parameters

|Name|Required/Optional|Data type|Description|
|:---|:----------------|:--------|:----------|
|_Direction_ | Required |**[XlErrorBarDirection](excel.xlerrorbardirection.md)** |The error bar direction. Can be one of these **XlErrorBarDirection** constants: **xlX** (can only be used with scatter charts) or **xlY** (default). |
|_Include_ |Required | **[XlErrorBarInclude](excel.xlerrorbarinclude.md)**|The error bar parts to be included. Can be one of these **XlErrorBarInclude** constants: **xlErrorBarIncludeBoth** (default), **xlErrorBarIncludeMinusValues**, **xlErrorBarIncludeNone**, or **xlErrorBarIncludePlusValues**. |
|_Type_ |Required |**[XlErrorBarType](excel.xlerrorbartype.md)**|The error bar type. Can be one of these **XlErrorBarType** constants: **xlErrorBarTypeCustom**, **xlErrorBarTypeFixedValue**, **xlErrorBarTypePercent**, **xlErrorBarTypeStDev**, or **xlErrorBarTypeStError**.|
|_Amount_ |Optional |**Variant**|The error amount. Used for only the positive error amount when _Type_ is **xlErrorBarTypeCustom**.|
|_MinusValues_ |Optional |**Variant**|The negative error amount when _Type_ is **xlErrorBarTypeCustom**.|

## Example

This example applies standard error bars in the Y direction for series one. The error bars are applied in the positive and negative directions. The example should be run on a 2D line chart.

```vb
myChart.SeriesCollection(1).ErrorBar _ 
 Direction:=xlY, Include:=xlErrorBarIncludeBoth, _ 
 Type:=xlErrorBarTypeStError
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]