---
title: Table.GetRowCount method (Outlook)
keywords: vbaol11.chm2232
f1_keywords:
- vbaol11.chm2232
ms.prod: outlook
api_name:
- Outlook.Table.GetRowCount
ms.assetid: 06014c43-700a-8502-bad7-b3f93a22e870
ms.date: 06/08/2017
localization_priority: Normal
---


# Table.GetRowCount method (Outlook)

Obtains the number of rows in the  **[Table](Outlook.Table.md)**.


## Syntax

_expression_. `GetRowCount`

_expression_ A variable that represents a [Table](Outlook.Table.md) object.


## Return value

A Long value that represents the number of rows in the Table.


## Remarks

 **GetRowCount** on a large table will result in a performance impact. Due to MAPI restrictions (for example, memory constraints for large tables, simultaneous operations on the **Table**), **GetRowCount** may not be able to determine the number of rows in the **Table**, or it may only return an approximate row count. In these cases, **GetRowCount** will return an error. You should use appropriate error detection for **GetRowCount** to determine if the call returns an error.


## See also


[Table Object](Outlook.Table.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]