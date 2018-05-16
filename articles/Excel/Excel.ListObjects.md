---
title: ListObjects Object (Excel)
keywords: vbaxl10.chm731072
f1_keywords:
- vbaxl10.chm731072
ms.prod: excel
api_name:
- Excel.ListObjects
ms.assetid: 3a888055-1ed0-d37d-0586-ced999dc1c42
ms.date: 06/08/2017
---


# ListObjects Object (Excel)

A collection of all the  **[ListObject](Excel.ListObject.md)** objects on a worksheet. Each **ListObject** object represents a table in the worksheet.


## Remarks

Use the  **[ListObjects](Excel.Worksheet.ListObjects.md)** property of the[Worksheet](Excel.Worksheet.md) object to return the **ListObjects** collection.


## Example

 The following example creates a new **ListObjects** collection which represents all the tables in a worksheet.


```
Set myWorksheetLists = Worksheets(1).ListObjects
```


## Methods



|**Name**|
|:-----|
|[Add](Excel.ListObjects.Add.md)|

## Properties



|**Name**|
|:-----|
|[Application](Excel.ListObjects.Application.md)|
|[Count](Excel.ListObjects.Count.md)|
|[Creator](Excel.ListObjects.Creator.md)|
|[Item](Excel.ListObjects.Item.md)|
|[Parent](Excel.ListObjects.Parent.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
