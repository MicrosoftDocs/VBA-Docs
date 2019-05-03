---
title: QueryTable.Sort property (Excel)
keywords: vbaxl10.chm518139
f1_keywords:
- vbaxl10.chm518139
ms.prod: excel
api_name:
- Excel.QueryTable.Sort
ms.assetid: 92f268ef-507f-a565-be42-abea73c381a2
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.Sort property (Excel)

Returns the sort criteria for the query table range. Read-only.


## Syntax

_expression_.**Sort**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

If you import data by using the user interface, data from web queries or text queries is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from web queries or text queries must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

You can use the **[QueryTable](Excel.ListObject.QueryTable.md)** property of the **ListObject** to access the **Sort** property.


## Example

This example refreshes the query table and gets the sort criteria.

```vb
QueryTable.Refresh 
 
Dim var As Sort 
Set var = QueryTable.Sort
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]