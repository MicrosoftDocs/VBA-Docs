---
title: QueryTable.RowNumbers property (Excel)
keywords: vbaxl10.chm518075
f1_keywords:
- vbaxl10.chm518075
ms.prod: excel
api_name:
- Excel.QueryTable.RowNumbers
ms.assetid: e0e91e2a-f7b6-ef5b-8046-9e93a51395db
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.RowNumbers property (Excel)

**True** if row numbers are added as the first column of the specified query table. Read/write **Boolean**.


## Syntax

_expression_.**RowNumbers**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

Setting this property to **True** doesn't immediately cause row numbers to appear. The row numbers appear the next time the query table is refreshed, and they're reconfigured every time the query table is refreshed.

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

You can use the **[QueryTable](Excel.ListObject.QueryTable.md)** property of the **ListObject** to access the **RowNumbers** property.


## Example

This example adds row numbers and field names to the query table.

```vb
With Worksheets(1).QueryTables("ExternalData1") 
 .RowNumbers = True 
 .FieldNames = True 
 .Refresh 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]