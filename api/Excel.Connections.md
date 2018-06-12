---
title: Connections Object (Excel)
keywords: vbaxl10.chm775072
f1_keywords:
- vbaxl10.chm775072
ms.prod: excel
api_name:
- Excel.Connections
ms.assetid: 3320b1cc-2f9d-805e-e506-27164b38d413
ms.date: 06/08/2017
---


# Connections Object (Excel)

A collection of Connection objects for the specified workbook.


## Example

The following example shows how to add a connection to a workbook from an existing file.


```
ActiveWorkbook.Connections.AddFromFile _ 
 "C:\Documents and Settings\myComputer\My Documents\My Data Sources\Northwind 2007 Customers.odc"
```


## Methods



|**Name**|
|:-----|
|[Add2](Excel.Connections.Add.md)|
|[AddFromFile](Excel.Connections.AddFromFile.md)|
|[Item](Excel.Connections.Item.md)|

## Properties



|**Name**|
|:-----|
|[Application](Excel.Connections.Application.md)|
|[Count](Excel.Connections.Count.md)|
|[Creator](Excel.Connections.Creator.md)|
|[Parent](Excel.Connections.Parent.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
