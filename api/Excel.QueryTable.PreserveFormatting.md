---
title: QueryTable.PreserveFormatting property (Excel)
keywords: vbaxl10.chm518111
f1_keywords:
- vbaxl10.chm518111
ms.prod: excel
api_name:
- Excel.QueryTable.PreserveFormatting
ms.assetid: 0be15116-ff1a-9b39-ae59-46c2d9383f0b
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.PreserveFormatting property (Excel)

**True** if any formatting common to the first five rows of data are applied to new rows of data in the query table. Unused cells aren't formatted. The property is **False** if the last AutoFormat applied to the query table is applied to new rows of data. The default value is **True**.


## Syntax

_expression_.**PreserveFormatting**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

For database query tables, the default formatting setting is the **xlSimple** [constant](excel.constants.md).

The new AutoFormat style is applied to the query table when the table is refreshed. The AutoFormat is reset to **None** whenever **PreserveFormatting** is set to **False**. As a result, any AutoFormat that's set before **PreserveFormatting** is set to **False** and before the query table is refreshed doesn't take effect, and the resulting query table has no formatting applied to it.

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

You can use the **[QueryTable](Excel.ListObject.QueryTable.md)** property of the **ListObject** to access the **PreserveFormatting** property.


## Example

This example preserves the formatting of the first PivotTable report on worksheet one.

```vb
Worksheets(1).PivotTables("Pivot1").PreserveFormatting = True
```

<br/>

This example demonstrates how setting **PreserveFormatting** to **False** causes the AutoFormat to be set to **xlRangeAutoFormatNone** instead of the specified **xlRangeAutoFormatColor1** format.

```vb
With Workbooks(1).Worksheets(1).QueryTables(1) 
 .Range.AutoFormat = xlRangeAutoFormatColor1 
 .PreserveFormatting = False 
 .Refresh 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
