---
title: Connections object (Excel)
keywords: vbaxl10.chm775072
f1_keywords:
- vbaxl10.chm775072
ms.prod: excel
api_name:
- Excel.Connections
ms.assetid: 3320b1cc-2f9d-805e-e506-27164b38d413
ms.date: 03/29/2019
localization_priority: Normal
---


# Connections object (Excel)

A collection of **[WorkbookConnection](Excel.WorkbookConnection.md)** objects for the specified workbook.


## Example

The following example shows how to add a connection to a workbook from an existing file.

```vb
ActiveWorkbook.Connections.AddFromFile _ 
 "C:\Documents and Settings\myComputer\My Documents\My Data Sources\Northwind 2007 Customers.odc"

```


## Methods

- [Add](Excel.Connections.Add.md)
- [AddFromFile](Excel.Connections.AddFromFile.md)
- [Item](Excel.Connections.Item.md)

## Properties

- [Application](Excel.Connections.Application.md)
- [Count](Excel.Connections.Count.md)
- [Creator](Excel.Connections.Creator.md)
- [Parent](Excel.Connections.Parent.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
