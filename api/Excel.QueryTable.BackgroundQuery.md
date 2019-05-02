---
title: QueryTable.BackgroundQuery property (Excel)
keywords: vbaxl10.chm518081
f1_keywords:
- vbaxl10.chm518081
ms.prod: excel
api_name:
- Excel.QueryTable.BackgroundQuery
ms.assetid: d3fd1d37-4956-7fda-accc-25eedf5188c0
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.BackgroundQuery property (Excel)

**True** if queries for the query table are performed asynchronously (in the background). Read/write **Boolean**.


## Syntax

_expression_.**BackgroundQuery**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

For OLAP data sources, this property is read-only and always returns **False**.

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

You can use the **[QueryTable](Excel.ListObject.QueryTable.md)** property of the **ListObject** to access the **BackgroundQuery** property.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
