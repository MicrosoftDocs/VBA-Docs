---
title: Table.Columns property (Outlook)
keywords: vbaol11.chm2236
f1_keywords:
- vbaol11.chm2236
ms.prod: outlook
api_name:
- Outlook.Table.Columns
ms.assetid: 57005ab1-ad49-296d-5b34-24dfd8f0987f
ms.date: 06/08/2017
localization_priority: Normal
---


# Table.Columns property (Outlook)

Returns a  **[Columns](Outlook.Columns.md)** collection object that contains the columns defined for the **[Table](Outlook.Table.md)**. Read-only.


## Syntax

_expression_.**Columns**

_expression_ A variable that represents a [Table](Outlook.Table.md) object.


## Remarks

The  **Columns** collection object is the default member of the **Table** object.

While rows in a  **Table** correspond to items in the parent **[Folder](Outlook.Folder.md)** or **[Search](Outlook.Search.md)** object of the **Table**, **Columns** in a **Table** correspond to the properties of these items. Default columns are defined for all folders depending on the parent folder of the **Table** object. For example, the default properties for the Inbox are: **EntryID**, **Subject**, **CreationTime**, **LastModificationTime**, and **MessageClass**. For more information on default properties for a **Table**, see [Default Properties Displayed in a Table Object](../outlook/How-to/Search-and-Filter/default-properties-displayed-in-a-table-object.md).

To add  **[Column](Outlook.Column.md)** objects to the **Columns** collection of a **Table**, use **[Columns.Add](Outlook.Columns.Add.md)**. To remove the default column set, use **[Columns.RemoveAll](Outlook.Columns.RemoveAll.md)**. For more information on adjusting columns of a **Table**, see [Adding Columns to a Table Object](../outlook/How-to/Search-and-Filter/adding-columns-to-a-table-object.md).


## See also


[Table Object](Outlook.Table.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]