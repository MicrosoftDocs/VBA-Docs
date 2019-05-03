---
title: QueryTable.TextFileStartRow property (Excel)
keywords: vbaxl10.chm518099
f1_keywords:
- vbaxl10.chm518099
ms.prod: excel
api_name:
- Excel.QueryTable.TextFileStartRow
ms.assetid: 91b774d8-cf7b-354d-510e-a8561076532c
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.TextFileStartRow property (Excel)

Returns or sets the row number at which text parsing will begin when you import a text file into a query table. Valid values are integers from 1 through 32767. The default value is 1. Read/write **Long**.


## Syntax

_expression_.**TextFileStartRow**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

Use this property only when your query table is based on data from a text file (with the **[QueryType](Excel.QueryTable.QueryType.md)** property set to **xlTextImport**).

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

The **TextFileStartRow** property applies only to **QueryTable** objects.


## Example

This example sets row 5 as the starting row for text parsing in the query table on the first worksheet in the first workbook, and then refreshes the query table.

```vb
Set shFirstQtr = Workbooks(1).Worksheets(1) 
Set qtQtrResults = shFirstQtr.QueryTables _ 
 .Add(Connection := "TEXT;C:\My Documents\19980331.txt", _ 
 Destination := shFirstQtr.Cells(1,1)) 
With qtQtrResults 
 .TextFileParseType = xlDelimited 
 .TextFileStartRow = 5 
 .TextFileTabDelimiter = True 
 .Refresh 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]