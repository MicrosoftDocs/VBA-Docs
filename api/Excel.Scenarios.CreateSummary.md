---
title: Scenarios.CreateSummary method (Excel)
keywords: vbaxl10.chm362075
f1_keywords:
- vbaxl10.chm362075
ms.prod: excel
api_name:
- Excel.Scenarios.CreateSummary
ms.assetid: b223ad02-cd11-7adc-2144-5c6dd1683427
ms.date: 05/11/2019
localization_priority: Normal
---


# Scenarios.CreateSummary method (Excel)

Creates a new worksheet that contains a summary report for the scenarios on the specified worksheet. **Variant**.


## Syntax

_expression_.**CreateSummary** (_ReportType_, _ResultCells_)

_expression_ A variable that represents a **[Scenarios](Excel.Scenarios.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ReportType_|Optional| **[XlSummaryReportType](Excel.XlSummaryReportType.md)**|Specifies whether the summary report is a PivotTable or a standard summary.|
| _ResultCells_|Optional| **Variant**|A **[Range](excel.range(object).md)** object that represents the result cells on the specified worksheet.<br/><br/>Normally, this range refers to one or more cells containing the formulas that depend on the changing cell values for your model; that is, the cells that show the results of a particular scenario. If this argument is omitted, there are no result cells included in the report.|

## Return value

Variant


## Example

This example creates a summary of the scenarios on Sheet1, with result cells in the range C4:C9 on Sheet1.

```vb
Worksheets("Sheet1").Scenarios.CreateSummary _ 
 ResultCells := Worksheets("Sheet1").Range("C4:C9")
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]