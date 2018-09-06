---
title: QueryTable.ListObject Property (Excel)
keywords: vbaxl10.chm518136
f1_keywords:
- vbaxl10.chm518136
ms.prod: excel
api_name:
- Excel.QueryTable.ListObject
ms.assetid: a302d0ac-7084-ba20-4e01-fe5e93bac307
ms.date: 06/08/2017
---


# QueryTable.ListObject Property (Excel)

Returns a  **[ListObject](Excel.ListObject.md)** object for the **[QueryTable](Excel.QueryTable.md)** object. Read-only **ListObject** object.


## Syntax

 _expression_. `ListObject`

 _expression_ A variable that represents a [QueryTable](Excel.QueryTable.md) object.


## Remarks

If you import data using the user interface, data from a Web query or a text query is imported as a  **[QueryTable](Excel.QueryTable.md)** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data using the object model, data from a Web query or a text query must be imported as a  **QueryTable** , while all other external data can be imported as either a **ListObject** or a **QueryTable** .

The  **ListObject** property applies only to **ListObject** objects.


## See also


[QueryTable Object](Excel.QueryTable.md)

