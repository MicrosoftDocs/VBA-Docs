---
title: Search.GetTable method (Outlook)
keywords: vbaol11.chm2261
f1_keywords:
- vbaol11.chm2261
ms.prod: outlook
api_name:
- Outlook.Search.GetTable
ms.assetid: 3aba6b77-73a3-9620-9c18-b2e03c7b63bc
ms.date: 06/08/2017
localization_priority: Normal
---


# Search.GetTable method (Outlook)

Obtains a **[Table](Outlook.Table.md)** object that contains items filtered by the _Filter_ parameter in a preceding **[Application.AdvancedSearch](Outlook.Application.AdvancedSearch.md)** method call.


## Syntax

_expression_. `GetTable`

_expression_ A variable that represents a [Search](Outlook.Search.md) object.


## Return value

A **Table** that contains items that meet the criteria specified by the _Filter_ parameter in a preceding **Application.AdvancedSearch** method call.


## Remarks

Unlike  **[Folder.GetTable](Outlook.Folder.GetTable.md)**, **Search.GetTable** does not accept a _Filter_ parameter. The filter for the **Table** is determined by **[Search.Filter](Outlook.Search.Filter.md)**. Since **Search.Filter** is a read-only property, the _Filter_ parameter for the **Application.AdvancedSearch** method establishes the filter for the **Table** object returned by **Search.GetTable**.

The  _Filter_ parameter supplied to **Application.AdvancedSearch** must be a DASL query. The filter for **AdvancedSearch** will not accept a JET query. Do not prefix a DASL query for **AdvancedSearch** with "@SQL=". If you add the "@SQL=" prefix, your query will raise an error. For more information on filters, see [Filtering Items](../outlook/How-to/Search-and-Filter/filtering-items.md).

 **Search.GetTable** returns a **Table** with the default column set for the folder type of the parent **Folder**. To modify the default column set, use the **[Add](Outlook.Columns.Add.md)**, **[Remove](Outlook.Columns.Remove.md)**, or **[RemoveAll](Outlook.Columns.RemoveAll.md)** methods of the **[Columns](Outlook.Columns.md)** collection object. For more information on default column sets, see [Default Properties Displayed in a Table Object](../outlook/How-to/Search-and-Filter/default-properties-displayed-in-a-table-object.md).

Unlike  **Folder.GetTable**, you cannot use **[Table.Restrict](Outlook.Table.Restrict.md)** to apply subsequent filters to a **Table** that is based on the **Search** object. Specify a new filter in **Application.AdvancedSearch** to re-apply a filter.


## See also


[Search Object](Outlook.Search.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]