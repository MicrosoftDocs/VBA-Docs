---
title: QueryTable.FillAdjacentFormulas property (Excel)
keywords: vbaxl10.chm518076
f1_keywords:
- vbaxl10.chm518076
ms.prod: excel
api_name:
- Excel.QueryTable.FillAdjacentFormulas
ms.assetid: 513a9218-a0b9-2bf6-ebac-1d9e7bb594df
ms.date: 05/03/2019
localization_priority: Normal
---


# QueryTable.FillAdjacentFormulas property (Excel)

**True** if formulas to the right of the specified query table are automatically updated whenever the query table is refreshed. Read/write **Boolean**.


## Syntax

_expression_.**FillAdjacentFormulas**

_expression_ A variable that represents a **[QueryTable](Excel.QueryTable.md)** object.


## Remarks

If you import data by using the user interface, data from a web query or a text query is imported as a **QueryTable** object, while all other external data is imported as a **[ListObject](Excel.ListObject.md)** object.

If you import data by using the object model, data from a web query or a text query must be imported as a **QueryTable**, while all other external data can be imported as either a **ListObject** or a **QueryTable**.

The **FillAdjacentFormulas** property applies only to **QueryTable** objects.


## Example

This example sets query table one so that formulas to the right of it are automatically updated whenever the query table is refreshed.

```vb
Sheets("sheet1").QueryTables(1).FillAdjacentFormulas = True
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]