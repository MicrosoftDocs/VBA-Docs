---
title: AddChartAutoFormat method (Excel Graph)
keywords: vbagr10.chm65752
f1_keywords:
- vbagr10.chm65752
ms.prod: excel
api_name:
- Excel.AddChartAutoFormat
ms.assetid: a75fe1dd-72f5-526c-a9b4-a309825e84d7
ms.date: 04/06/2019
localization_priority: Normal
---


# AddChartAutoFormat method (Excel Graph)

Adds a custom chart autoformat to the list of available chart autoformats.

## Syntax

_expression_.**AddChartAutoFormat** (_Name_, _Description_)

_expression_ Required. An expression that returns an **[Application](excel.application-graph-object.md)** object.

## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Name_ |Required |**String** |The name of the autoformat.|
|_Description_ | Optional | **String** |A description of the custom autoformat. |

## Example

This example adds a new autoformat.

```vb
myChart.Application.AddChartAutoFormat _ 
 Name:="Presentation Chart"
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]