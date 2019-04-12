---
title: CalculatedFields object (Excel)
keywords: vbaxl10.chm243072
f1_keywords:
- vbaxl10.chm243072
ms.prod: excel
api_name:
- Excel.CalculatedFields
ms.assetid: 6db4c889-f097-9a66-abc6-28f7f54f0478
ms.date: 03/29/2019
localization_priority: Normal
---


# CalculatedFields object (Excel)

A collection of **[PivotField](Excel.PivotField.md)** objects that represents all the calculated fields in the specified PivotTable report.


## Remarks

A report that contains Revenue and Expense fields could have a calculated field named Profit defined as the amount in the Revenue field minus the amount in the Expense field.

For OLAP data sources, you cannot set this collection, and it always returns **Nothing**.

Use the **[CalculatedFields](Excel.PivotTable.CalculatedFields.md)** method of the **PivotTable** object to return the **CalculatedFields** collection.

Use **CalculatedFields** (_index_), where _index_ is the specified field's name or index number, to return a single **PivotField** object from the **CalculatedFields** collection.


## Example

The following example deletes the calculated fields from the PivotTable report named **Pivot1**.

```vb
For Each fld in _ 
 Worksheets(1).PivotTables("Pivot1").CalculatedFields 
 fld.Delete 
Next
```


## Methods

- [Add](Excel.CalculatedFields.Add.md)
- [Item](Excel.CalculatedFields.Item.md)

## Properties

- [Application](Excel.CalculatedFields.Application.md)
- [Count](Excel.CalculatedFields.Count.md)
- [Creator](Excel.CalculatedFields.Creator.md)
- [Parent](Excel.CalculatedFields.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]