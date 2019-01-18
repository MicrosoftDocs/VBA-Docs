---
title: QueryTable.ResultRange property (Excel)
keywords: vbaxl10.chm518090
f1_keywords:
- vbaxl10.chm518090
ms.prod: excel
api_name:
- Excel.QueryTable.ResultRange
ms.assetid: 7d7bde05-0e46-a282-dbdc-b2f5edcc2000
ms.date: 06/08/2017
localization_priority: Priority
---


# QueryTable.ResultRange property (Excel)

Returns a  **[Range](Excel.Range(object).md)** object that represents the area of the worksheet occupied by the specified query table. Read-only.


## Syntax

_expression_. `ResultRange`

_expression_ A variable that represents a [QueryTable](Excel.QueryTable.md) object.


## Remarks

The range doesn't include the field name row or the row number column.

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](Excel.QueryTable.md)** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable**.

You can use the  **[QueryTable](Excel.ListObject.QueryTable.md)** property of the **ListObject** to access the **ResultRange** property.


## Example

This example sums the data in the first column of query table one. The sum of the first column is displayed below the data range.


```vb
Set c1 = Sheets("sheet1").QueryTables(1).ResultRange.Columns(1) 
c1.Name = "Column1" 
c1.End(xlDown).Offset(2, 0).Formula = "=sum(Column1)"
```


## See also


[QueryTable Object](Excel.QueryTable.md)

