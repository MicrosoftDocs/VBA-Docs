---
title: Adding Columns to a Table Object
ms.prod: outlook
ms.assetid: c1d652ef-8082-70f3-1216-d39e976e6b21
ms.date: 06/08/2017
localization_priority: Normal
---


# Adding Columns to a Table Object

This topic describes how to add columns to a **[Table](../../../api/Outlook.Table.md)** object.

To obtain an initial **Table** object, use **[Folder.GetTable](../../../api/Outlook.Folder.GetTable.md)** or **[Search.GetTable](../../../api/Outlook.Search.GetTable.md)**. The returned **Table** object always contains a default set of properties depending on the folder type of the parent folder. If you want to change the columns in a **Table**, start with the **Table** object returned from a prior **GetTable** call. Use **[Table.Columns](../../../api/Outlook.Table.Columns.md)** to obtain the **[Columns](../../../api/Outlook.Columns.md)** object, and call **[Columns.Add](../../../api/Outlook.Columns.Add.md)**, **[Columns.Remove](../../../api/Outlook.Columns.Remove.md)**, or **[Columns.RemoveAll](../../../api/Outlook.Columns.RemoveAll.md)**. As a result of the call on the **Columns** object, the parent **Table** object is updated.

 **Note** Each of these calls on the **Columns** object adjusts the columns in the parent **Table**. The rows in the **Table** however remain the same as before the call. You do not call **GetTable** subsequently to obtain an updated **Table**. **GetTable** always returns a **Table** with the default set of columns for that folder type.

Since a folder can contain heterogeneous items (for example, the Deleted Items folder), you can use **Columns.Add** to add columns that do not apply to all rows in that Table. In such cases, **[Row.Item](../../../api/Outlook.Row.Item.md)** would return an error indicating that an object could not be found for the row at the specific column. Consequently, before you access other column values in a **Table**, you should first check for the **MessageClass** of a row (by calling `Row.Item("MessageClass")`) to determine which columns in the **Table** apply to that row.

 **Note** Since the **Item** method is the default method for the **[Row](../../../api/Outlook.Row.md)** object, `Row.Item("MessageClass")` is equivalent to `Row("MessageClass")`.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]