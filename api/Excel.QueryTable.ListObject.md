---
title: QueryTable.ListObject property (Excel)
keywords: vbaxl10.chm518136
f1_keywords:
- vbaxl10.chm518136
ms.prod: excel
api_name:
- Excel.QueryTable.ListObject
ms.assetid: a302d0ac-7084-ba20-4e01-fe5e93bac307
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.ListObject property (Excel)

Returns a **[ListObject](Excel.ListObject.md)** object for the **QueryTable** object. Read-only **ListObject** object.


## Syntax

_expression_.**ListObject**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **ListObject** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

The **ListObject** property applies only to **ListObject** objects.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]