---
title: PivotTable.ConvertToFormulas method (Excel)
keywords: vbaxl10.chm235177
f1_keywords:
- vbaxl10.chm235177
ms.prod: excel
api_name:
- Excel.PivotTable.ConvertToFormulas
ms.assetid: 8646696c-47c0-3851-4310-5e5368475266
ms.date: 05/08/2019
localization_priority: Normal
---


# PivotTable.ConvertToFormulas method (Excel)

The **ConvertToFormulas** method is used for converting a PivotTable to cube formulas. Read/write **Boolean**.


## Syntax

_expression_.**ConvertToFormulas** (_ConvertFilters_)

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ConvertFilters_|Required| **Boolean**|Contains **True** or **False** to indicate the state of the **ReportFilter** area.|

## Remarks

The argument controls whether or not to convert the **ReportFilter** area of the PivotTable.


## Example

In the following example, the **ReportFilter** area is not converted.

```vb
Sub ConvertToCubeFormulas() 
 ActiveSheet.PivotTables("PivotTable1").ConvertToFormulas False 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]