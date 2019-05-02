---
title: QueryTable.RobustConnect property (Excel)
keywords: vbaxl10.chm518133
f1_keywords:
- vbaxl10.chm518133
ms.prod: excel
api_name:
- Excel.QueryTable.RobustConnect
ms.assetid: ad180446-82d7-7b5b-59a2-b0de299ae934
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.RobustConnect property (Excel)

Returns or sets how the query table connects to its data source. Read/write **[XlRobustConnect](Excel.XlRobustConnect.md)**.


## Syntax

_expression_.**RobustConnect**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

You can use the **[QueryTable](Excel.ListObject.QueryTable.md)** property of the **ListObject** to access the **RobustConnect** property.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]