---
title: QueryTable.Parent property (Excel)
keywords: vbaxl10.chm517075
f1_keywords:
- vbaxl10.chm517075
ms.prod: excel
api_name:
- Excel.QueryTable.Parent
ms.assetid: 6cf47be7-5e4a-31d0-0c11-e9506c052ecf
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.Parent property (Excel)

Returns the parent object for the specified object. Read-only.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

You can use the **[QueryTable](Excel.ListObject.QueryTable.md)** property of the **ListObject** to access the **Parent** property.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]