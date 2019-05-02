---
title: QueryTable.EnableRefresh property (Excel)
keywords: vbaxl10.chm518084
f1_keywords:
- vbaxl10.chm518084
ms.prod: excel
api_name:
- Excel.QueryTable.EnableRefresh
ms.assetid: 79a0b628-b90d-1795-830f-e05bc6043517
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.EnableRefresh property (Excel)

**True** if the PivotTable cache or query table can be refreshed by the user. The default value is **True**. Read/write **Boolean**.


## Syntax

_expression_.**EnableRefresh**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

The **[RefreshOnFileOpen](Excel.QueryTable.RefreshOnFileOpen.md)** property is ignored if the **EnableRefresh** property is set to **False**.

For OLAP data sources, setting this property to **False** disables updates.

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

You can use the **[QueryTable](Excel.ListObject.QueryTable.md)** property of the **ListObject** to access the **EnableRefresh** property.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]