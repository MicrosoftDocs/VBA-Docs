---
title: PivotCache.RefreshOnFileOpen property (Excel)
keywords: vbaxl10.chm227083
f1_keywords:
- vbaxl10.chm227083
ms.prod: excel
api_name:
- Excel.PivotCache.RefreshOnFileOpen
ms.assetid: aed513aa-b752-8b6e-0d6d-6fddab46df18
ms.date: 05/03/2019
localization_priority: Normal
---


# PivotCache.RefreshOnFileOpen property (Excel)

**True** if the PivotTable cache is automatically updated each time the workbook is opened. The default value is **False**. Read/write **Boolean**.


## Syntax

_expression_.**RefreshOnFileOpen**

_expression_ A variable that represents a **[PivotCache](Excel.PivotCache.md)** object.


## Remarks

Query tables and PivotTable reports are not automatically refreshed when you open the workbook by using the **[Open](Excel.Workbooks.Open.md)** method of the **Workbooks** object in Visual Basic. Use the **[Refresh](Excel.PivotCache.Refresh.md)** method to refresh the data after the workbook is open.


## Example

This example causes the PivotTable cache to automatically update each time the workbook is opened.

```vb
ActiveWorkbook.PivotCaches(1).RefreshOnFileOpen = True
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]