---
title: QueryTable.Recordset property (Excel)
keywords: vbaxl10.chm518094
f1_keywords:
- vbaxl10.chm518094
ms.prod: excel
api_name:
- Excel.QueryTable.Recordset
ms.assetid: d9f4190e-c43c-5fe5-113d-18c8efcc2a27
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.Recordset property (Excel)

Returns or sets a **Recordset** object that's used as the data source for the specified query table. Read/write.


## Syntax

_expression_.**Recordset**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

If this property is used to overwrite an existing recordset, the change takes effect when the **[Refresh](Excel.QueryTable.Refresh.md)** method is run.

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

You can use the **[QueryTable](Excel.ListObject.QueryTable.md)** property of the **ListObject** to access the **RecordSet** property.


## Example

This example changes the **Recordset** object used with the first query table on the first worksheet and then refreshes the query table.

```vb
With Worksheets(1).QueryTables(1) 
 .Recordset = _ 
 Workbooks.OpenDatabase("c:\Nwind.mdb") _ 
 .OpenRecordset("employees") 
 .Refresh 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]