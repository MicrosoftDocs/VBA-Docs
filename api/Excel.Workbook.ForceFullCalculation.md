---
title: Workbook.ForceFullCalculation property (Excel)
keywords: vbaxl10.chm199264
f1_keywords:
- vbaxl10.chm199264
api_name:
- Excel.Workbook.ForceFullCalculation
ms.assetid: 76f46d18-79e3-9828-d126-e221ae1a8157
ms.date: 05/29/2019
ms.localizationpriority: medium
---


# Workbook.ForceFullCalculation property (Excel)

Returns or sets the specified workbook to forced calculation mode. Read/write.


## Syntax

_expression_.**ForceFullCalculation**

_expression_ An expression that returns a **[Workbook](Excel.Workbook.md)** object.


## Return value

**Boolean**


## Remarks

If the workbook is in the forced calculation mode, dependencies are ignored and all worksheets are marked to calculate fully every time a calculation is triggered. This setting remains in effect until Excel is restarted.

Setting the **ForceFullCalculation** property to **True** will increase the calculation times for data tables in proportion to the size of the data table. Given an NxM data table, the calculation time will increase by about _base time_ x (_N_ x _M_) so that a 3x4 data table may take about 12 times as long to calculate if this property is set to **True**.


## Example

The following example sets the active workbook to forced calculation mode.

```vb
ActiveWorkbook.ForceFullCalculation = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]