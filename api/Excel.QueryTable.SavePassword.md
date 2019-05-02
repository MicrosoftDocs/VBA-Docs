---
title: QueryTable.SavePassword property (Excel)
keywords: vbaxl10.chm518085
f1_keywords:
- vbaxl10.chm518085
ms.prod: excel
api_name:
- Excel.QueryTable.SavePassword
ms.assetid: c17250b1-9f80-12ed-1cbf-9f54a05ebaf3
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.SavePassword property (Excel)

**True** if password information in an ODBC connection string is saved with the specified query. **False** if the password is removed. Read/write **Boolean**.


## Syntax

_expression_.**SavePassword**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

This property is used in both ODBC and OLE DB queries, and by both PivotTables and QueryTables.

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

You can use the **[QueryTable](Excel.ListObject.QueryTable.md)** property of the **ListObject** to access the **SavePassword** property.

This property is ignored if the **ListObject** is connected to a SharePoint list.


## Example

This example causes password information to be removed from the ODBC connection string whenever query table one is saved.

```vb
Worksheets(1).QueryTables(1).SavePassword = False
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]