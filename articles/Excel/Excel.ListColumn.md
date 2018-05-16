---
title: ListColumn Object (Excel)
keywords: vbaxl10.chm737072
f1_keywords:
- vbaxl10.chm737072
ms.prod: excel
api_name:
- Excel.ListColumn
ms.assetid: c2060e4a-2340-c606-f272-1e4dad6964d0
ms.date: 06/08/2017
---


# ListColumn Object (Excel)

Represents a column in a table.


## Remarks

 The **ListColumn** object is a member of the **[ListColumns](Excel.ListColumns.md)** collection. The **ListColumns** collection contains all the columns in a table ( **[ListObject](Excel.ListObject.md)** object).

Use the [ListColumns](Excel.ListObject.ListColumns.md) property of the **ListObject** object to return a **[ListColumns](Excel.ListColumns.md)** collection.


## Example

The following example adds a new  **ListColumn** object to the default **ListObject** object in the first worksheet of the active workbook. Because no position is specified, a new rightmost column is added.


```
Sub AddListColumn() 
 Dim wrksht As Worksheet 
 Dim objListCol As ListColumn 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListCol = wrksht.ListObjects(1).ListColumns.Add 
End Sub 

```


## Methods



|**Name**|
|:-----|
|[Delete](Excel.ListColumn.Delete.md)|

## Properties



|**Name**|
|:-----|
|[Application](Excel.ListColumn.Application.md)|
|[Creator](Excel.ListColumn.Creator.md)|
|[DataBodyRange](Excel.ListColumn.DataBodyRange.md)|
|[Index](Excel.ListColumn.Index.md)|
|[Name](Excel.ListColumn.Name.md)|
|[Parent](Excel.ListColumn.Parent.md)|
|[Range](Excel.ListColumn.Range.md)|
|[Total](Excel.ListColumn.Total.md)|
|[TotalsCalculation](Excel.ListColumn.TotalsCalculation.md)|
|[XPath](listcolumn-xpath-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
