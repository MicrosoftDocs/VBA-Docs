---
title: QueryTable.EnableEditing property (Excel)
keywords: vbaxl10.chm518097
f1_keywords:
- vbaxl10.chm518097
ms.prod: excel
api_name:
- Excel.QueryTable.EnableEditing
ms.assetid: c8297f41-56fa-4d8c-6633-bbda0deb6257
ms.date: 06/08/2017
localization_priority: Normal
---


# QueryTable.EnableEditing property (Excel)

 **True** if the user can edit the specified query table. **False** if the user can only refresh the query table. Read/write **Boolean**.


## Syntax

_expression_. `EnableEditing`

_expression_ A variable that represents a [QueryTable](Excel.QueryTable.md) object.


## Remarks

This example sets query table one so that the user cannot edit it.

If you import data using the user interface, data from a web query or a text query is imported as a  **[QueryTable](Excel.QueryTable.md)** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data using the object model, data from a web query or a text query must be imported as a  **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

You can use the  **[QueryTable](Excel.ListObject.QueryTable.md)** property of the **ListObject** to access the **EnableEditing** property.


## Example


```vb
Worksheets(1).QueryTables(1).EnableEditing = False
```


## See also


[QueryTable Object](Excel.QueryTable.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]