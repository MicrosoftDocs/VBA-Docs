---
title: QueryTable.ResetTimer method (Excel)
keywords: vbaxl10.chm518121
f1_keywords:
- vbaxl10.chm518121
ms.prod: excel
api_name:
- Excel.QueryTable.ResetTimer
ms.assetid: 9e8c9d26-fe11-90f7-6073-c8ff5be3042d
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.ResetTimer method (Excel)

Resets the refresh timer for the specified query table or PivotTable report to the last interval that you set by using the **[RefreshPeriod](Excel.QueryTable.RefreshPeriod.md)** property.


## Syntax

_expression_.**ResetTimer**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Example

This example resets the refresh timer for the first query table on the active worksheet.

```vb
ActiveSheet.QueryTables(1).ResetTimer
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]