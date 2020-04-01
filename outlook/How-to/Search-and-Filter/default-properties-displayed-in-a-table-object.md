---
title: Default Properties Displayed in a Table Object
ms.prod: outlook
ms.assetid: 649c64f3-2d1e-23f1-bf13-3368da79e62b
ms.date: 06/08/2017
localization_priority: Normal
---


# Default Properties Displayed in a Table Object

A **[Table](../../../api/Outlook.Table.md)** contains rows and columns, where rows represent items in a folder, and columns represent item properties. When you call **[Folder.GetTable](../../../api/Outlook.Folder.GetTable.md)**, you obtain a **Table** object that has a small pre-defined set of columns corresponding to properties common to default items for that type of folder. Similarly, when you call **[Search.GetTable](../../../api/Outlook.Search.GetTable.md)**, you obtain a **Table** that has columns corresponding to properties common to default items for all folder types. The pre-defined sets of properties are explicit built-in properties. The small size of these sets allows the **GetTable** call to perform relatively efficiently.

The following tables list the set of properties returned by **GetTable** for each type of folder or a search folder. Properties are stored as a 1-based array in a **[Columns](../../../api/Outlook.Columns.md)** object.

## Columns for all Folder Types

The following table shows the properties that are returned as default columns in a **Table** for any folder, including a search folder, Inbox, Sent Items, Deleted Items, Journal, and Notes:

| **Column**| **Description**|
|:-----|:-----|
|1| **EntryID**|
|2| **Subject**|
|3| **CreationTime**|
|4| **LastModificationTime**|
|5| **MessageClass**|

## Columns for the Calendar Folder

The following table shows the properties that are returned as default columns in a **Table** for the Calendar folder:

| **Column**| **Description**|
|:-----|:-----|
|1| **EntryID**|
|2| **Subject**|
|3| **CreationTime**|
|4| **LastModificationTime**|
|5| **MessageClass**|
|6| **Start**|
|7| **End**|
|8| **IsRecurring**|


## Columns for the Contacts Folder

The following table shows the properties that are returned as default columns in a **Table** for the Contacts folder:


| **Column**| **Description**|
|:-----|:-----|
|1| **EntryID**|
|2| **Subject**|
|3| **CreationTime**|
|4| **LastModificationTime**|
|5| **MessageClass**|
|6| **FirstName**|
|7| **LastName**|
|8| **CompanyName**|


## Columns for the Task Folder

The following table shows the properties that are returned as default columns in a **Table** for the Task folder:

| **Column**| **Description**|
|:-----|:-----|
|1| **EntryID**|
|2| **Subject**|
|3| **CreationTime**|
|4| **LastModificationTime**|
|5| **MessageClass**|
|6| **DueDate**|
|7| **PercentComplete**|
|8| **IsRecurring**|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]