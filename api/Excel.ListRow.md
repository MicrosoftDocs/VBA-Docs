---
title: ListRow object (Excel)
keywords: vbaxl10.chm741072
f1_keywords:
- vbaxl10.chm741072
ms.prod: excel
api_name:
- Excel.ListRow
ms.assetid: ba3e4215-14b6-3dca-82d0-0951f9f2fc3e
ms.date: 03/30/2019
localization_priority: Normal
---


# ListRow object (Excel)

Represents a row in a table. The **ListRow** object is a member of the **[ListRows](Excel.ListRows.md)** collection.


## Remarks

The **ListRows** collection contains all the rows in a list object.

Use the **[ListRows](Excel.ListObject.ListRows.md)** property of the **ListObject** object to return a **ListRows** collection.


## Example

The following example adds a new **ListRow** object to the default **ListObject** object in the first worksheet of the active workbook. Because no position is specified, a new row is added to the end of the table.

```vb
Dim wrksht As Worksheet 
Dim oListRow As ListRow 
 
Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
Set oListRow = wrksht.ListObjects(1).ListRows.Add 

```


## Methods

- [Delete](Excel.ListRow.Delete.md)

## Properties

- [Application](Excel.ListRow.Application.md)
- [Creator](Excel.ListRow.Creator.md)
- [Index](Excel.ListRow.Index.md)
- [Parent](Excel.ListRow.Parent.md)
- [Range](Excel.ListRow.Range.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
