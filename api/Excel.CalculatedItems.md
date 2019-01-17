---
title: CalculatedItems object (Excel)
keywords: vbaxl10.chm249072
f1_keywords:
- vbaxl10.chm249072
ms.prod: excel
api_name:
- Excel.CalculatedItems
ms.assetid: daad9732-6a20-d146-050e-da9e1c1e6f33
ms.date: 06/08/2017
localization_priority: Normal
---


# CalculatedItems object (Excel)

A collection of  **[PivotItem](Excel.PivotItem.md)** objects that represent all the calculated items in the specified PivotTable report.


## Remarks

A PivotTable report that contains January, February, and March items could have a calculated item named "FirstQuarter" defined as the sum of the amounts in January, February, and March.

Use the  **[CalculatedItems](Excel.PivotField.CalculatedItems.md)** method to return the **CalculatedItems** collection.

Use  **CalculatedFields** ( _index_ ), where _index_ is the name or index number of the field, to return a single **[PivotField](Excel.PivotField.md)** object from the **[CalculatedFields](Excel.CalculatedFields.md)** collection.


## Example

The following example creates a list of the calculated items in the first PivotTable report on worksheet one, along with their formulas.


```vb
Set pt = Worksheets(1).PivotTables(1) 
For Each ci In pt.PivotFields("Sales").CalculatedItems 
 r = r + 1 
 With Worksheets(2) 
 .Cells(r, 1).Value = ci.Name 
 .Cells(r, 2).Value = ci.Formula 
 End With 
Next
```


## See also



[Excel Object Model Reference](overview/Excel/object-model.md)

