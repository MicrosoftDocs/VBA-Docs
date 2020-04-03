---
title: Column object (Outlook)
keywords: vbaol11.chm3191
f1_keywords:
- vbaol11.chm3191
ms.prod: outlook
api_name:
- Outlook.Column
ms.assetid: b7eb6916-2d80-57c3-2077-47a2a4c73185
ms.date: 06/08/2017
localization_priority: Normal
---


# Column object (Outlook)

Represents a column of data in a  **[Table](Outlook.Table.md)** object.


## Remarks

A  **Table** is composed of rows and columns. It represents a read-only dynamic rowset of data in a **[Folder](Outlook.Folder.md)** or **[Search](Outlook.Search.md)** object. You can regard each row of a **Table** as an item in the folder, each column as a property of the item. By default, a **Table** contains only a subset of properties for items in the folder. This makes the **Table** an in-memory lightweight rowset that allows fast enumeration and filtering of items in the folder.

To obtain the value of a property (column) for a specific item (row) in a  **Table**, you can either use the **[Table.GetArray](Outlook.Table.GetArray.md)** method and index into the returned array, or use the **[Row.Item](Outlook.Row.Item.md)** method, specifying the **[Name](Outlook.Column.Name.md)** of the column.


## Properties



|Name|
|:-----|
|[Application](Outlook.Column.Application.md)|
|[Class](Outlook.Column.Class.md)|
|[Name](Outlook.Column.Name.md)|
|[Parent](Outlook.Column.Parent.md)|
|[Session](Outlook.Column.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]