---
title: ListColumns object (Excel)
keywords: vbaxl10.chm735072
f1_keywords:
- vbaxl10.chm735072
ms.prod: excel
api_name:
- Excel.ListColumns
ms.assetid: c1b8aff0-3049-df58-ce1f-0c5e4bddc467
ms.date: 06/08/2017
localization_priority: Priority
---


# ListColumns object (Excel)

A collection of all the  **[ListColumn](Excel.ListColumn.md)** objects in the specified **[ListObject](Excel.ListObject.md)** object.


## Remarks

 Each **ListColumn** object represents a column in the table.


 **Note**  A name for the column is automatically generated. You can change the name after the column has been added.


## Example

Use the  **[ListColumns](Excel.ListObject.ListColumns.md)** property of the [ListObject](Excel.ListObject.md) object to return the **[ListColumns](Excel.ListColumns.md)** collection. The following example adds a new column to the default **ListObject** object in the first worksheet of the workbook. Because no position is specified, a new rightmost column is added.


```vb
Set myNewColumn = Worksheets(1).ListObject(1).ListColumns.Add
```


## Methods



|Name|
|:-----|
|[Add](Excel.ListColumns.Add.md)|

## Properties



|Name|
|:-----|
|[Application](Excel.ListColumns.Application.md)|
|[Count](Excel.ListColumns.Count.md)|
|[Creator](Excel.ListColumns.Creator.md)|
|[Item](Excel.ListColumns.Item.md)|
|[Parent](Excel.ListColumns.Parent.md)|

## See also


[Excel Object Model Reference](overview/Excel/object-model.md)
