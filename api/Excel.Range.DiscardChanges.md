---
title: Range.DiscardChanges method (Excel)
keywords: vbaxl10.chm144254
f1_keywords:
- vbaxl10.chm144254
ms.prod: excel
api_name:
- Excel.Range.DiscardChanges
ms.assetid: adeee827-d680-59f3-0966-2c2ca60a59e1
ms.date: 05/10/2019
localization_priority: Normal
---


# Range.DiscardChanges method (Excel)

Discards all changes in the edited cells of the range.


## Syntax

_expression_.**DiscardChanges**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Remarks

For a range based on an OLAP data source, this method removes all values and formulas entered in the cells, and then runs an update operation to retrieve the latest values from the data source. It sets the data source value to **NULL** for all value cells that are edited, and also executes a **ROLLBACK TRANSACTION** statement against the OLAP server. 

For ranges based on non-OLAP data sources, this method simply clears all edited cells.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]