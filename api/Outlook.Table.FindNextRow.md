---
title: Table.FindNextRow method (Outlook)
keywords: vbaol11.chm2229
f1_keywords:
- vbaol11.chm2229
ms.prod: outlook
api_name:
- Outlook.Table.FindNextRow
ms.assetid: e09019ca-e4bb-2597-7b9e-a56c1b5fce6c
ms.date: 06/08/2017
localization_priority: Normal
---


# Table.FindNextRow method (Outlook)

Finds the next row in the  **[Table](Outlook.Table.md)** that meets the criteria specified in a preceding **[Table.FindRow](Outlook.Table.FindRow.md)**.


## Syntax

_expression_. `FindNextRow`

_expression_ A variable that represents a [Table](Outlook.Table.md) object.


## Return value

A **[Row](Outlook.Row.md)** object that represents the next row in the **Table** that meets the filter condition in the preceding call to **FindRow**. Returns **Null** (**Nothing** in Visual Basic) if **FindNextRow** cannot find another row that meets the criteria specified in **FindRow**. Also returns **Null** if **FindRow** has not been called before **FindNextRow**.


## Remarks

 **FindNextRow** finds the next row based on the row returned by the preceding **FindRow** or **FindNextRow**. It does not depend on the current row (as the current row may have been repositioned since the preceding **FindRow** or **FindNextRow**, for example, by **[Table.MoveToStart](Outlook.Table.MoveToStart.md)**).

If  **FindNextRow** finds a row, it will position the current row to that row. If it does not find another row, it will not reposition the current row.


## See also


[Table Object](Outlook.Table.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]