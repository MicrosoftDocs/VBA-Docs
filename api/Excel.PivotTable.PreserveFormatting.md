---
title: PivotTable.PreserveFormatting property (Excel)
keywords: vbaxl10.chm235122
f1_keywords:
- vbaxl10.chm235122
ms.prod: excel
api_name:
- Excel.PivotTable.PreserveFormatting
ms.assetid: d37d215a-b031-5a20-f302-471df3a3b2a2
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.PreserveFormatting property (Excel)

**True** if formatting is preserved when the report is refreshed or recalculated by operations such as pivoting, sorting, or changing page field items.

For query tables, this property is **True** if any formatting common to the first five rows of data are applied to new rows of data in the query table. Unused cells aren't formatted. 

The property is **False** if the last AutoFormat applied to the query table is applied to new rows of data. The default value is **True**.


## Syntax

_expression_.**PreserveFormatting**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Remarks

For database query tables, the default formatting setting is the **xlSimple** [constant](excel.constants.md).

The new AutoFormat style is applied to the query table when the table is refreshed. The AutoFormat is reset to **None** whenever **PreserveFormatting** is set to **False**. As a result, any AutoFormat that's set before **PreserveFormatting** is set to **False** and before the query table is refreshed doesn't take effect, and the resulting query table has no formatting applied to it.


## Example

This example preserves the formatting of the first PivotTable report on worksheet one.

```vb
Worksheets(1).PivotTables("Pivot1").PreserveFormatting = True
```

<br/>

This example demonstrates how setting **PreserveFormatting** to **False** causes the AutoFormat to be set to the **[XlRangeAutoFormat](excel.xlrangeautoformat.md)** value **xlRangeAutoFormatNone** instead of the specified **xlRangeAutoFormatColor1** format.

```vb
With Workbooks(1).Worksheets(1).QueryTables(1) 
 .Range.AutoFormat = xlRangeAutoFormatColor1 
 .PreserveFormatting = False 
 .Refresh 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]