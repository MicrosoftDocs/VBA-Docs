---
title: ListColumn object (Excel)
keywords: vbaxl10.chm737072
f1_keywords:
- vbaxl10.chm737072
ms.prod: excel
api_name:
- Excel.ListColumn
ms.assetid: c2060e4a-2340-c606-f272-1e4dad6964d0
ms.date: 03/30/2019
localization_priority: Normal
---


# ListColumn object (Excel)

Represents a column in a table.


## Remarks

The **ListColumn** object is a member of the **[ListColumns](Excel.ListColumns.md)** collection. The **ListColumns** collection contains all the columns in a table.

Use the **[ListColumns](Excel.ListObject.ListColumns.md)** property of the **ListObject** object to return a **ListColumns** collection.


## Example

The following example adds a new **ListColumn** object to the default **ListObject** object in the first worksheet of the active workbook. Because no position is specified, a new rightmost column is added.

```vb
Sub AddListColumn() 
 Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns.Add 
End Sub 

```


## Methods

- [Delete](Excel.ListColumn.Delete.md)

## Properties

- [Application](Excel.ListColumn.Application.md)
- [Creator](Excel.ListColumn.Creator.md)
- [DataBodyRange](Excel.ListColumn.DataBodyRange.md)
- [Index](Excel.ListColumn.Index.md)
- [Name](Excel.ListColumn.Name.md)
- [Parent](Excel.ListColumn.Parent.md)
- [Range](Excel.ListColumn.Range.md)
- [Total](Excel.ListColumn.Total.md)
- [TotalsCalculation](Excel.ListColumn.TotalsCalculation.md)
- [XPath](Excel.ListColumn.XPath.md)


## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
