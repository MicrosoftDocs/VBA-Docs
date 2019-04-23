---
title: CodeData object (Access)
keywords: vbaac10.chm12742
f1_keywords:
- vbaac10.chm12742
ms.prod: access
api_name:
- Access.CodeData
ms.assetid: fc207136-4d18-2c7d-ffe6-0e1ad7c2fc32
ms.date: 02/27/2019
localization_priority: Normal
---


# CodeData object (Access)

The **CodeData** object refers to objects stored within the code database by the source (server) application.


## Remarks

The **CodeData** object has several collections that contain specific object types within the code database. The following table lists the name of each collection defined by the database and the types of objects it contains.

<br/>

|Collections|Object type|Available in Access database|Available in Access Project (.adp)|
|:-----|:-----|:-----|:-----|
|**[AllTables](Access.AllTables.md)**|All tables|Yes|Yes|
|**[AllFunctions](Access.AllFunctions.md)**|All functions|No|Yes|
|**[AllQueries](Access.AllQueries.md)**|All queries |Yes|Yes|
|**[AllViews](Access.AllViews.md)**|All views |No|Yes|
|**[AllStoredProcedures](Access.AllStoredProcedures.md)**|All stored procedures |No|Yes|
|**[AllDatabaseDiagrams](Access.AllDatabaseDiagrams.md)**|All database diagrams |No|Yes|

> [!NOTE] 
> The collections in the preceding table contain all of the respective objects in the database regardless if they are opened or closed.

For example, an **AccessObject** representing a table is a member of the **AllTables** collection, which is a collection of **AccessObject** objects within the current database. Within the **AllTables** collection, individual tables are indexed beginning with zero. You can refer to an individual **AccessObject** object in the **AllTables** collection either by referring to the table by name, or by referring to its index within the collection. If you want to refer to a specific item in the **AllTables** collection, it's better to refer to it by name because the item's index may change. If the object name includes a space, the name must be surrounded by brackets ([ ]).

<br/>

|Syntax|Example|
|:-----|:-----|
|**AllTables**!_tablename_|AllTables!OrderTable|
|**AllTables**![_table name_]|AllTables![Order Table]|
|**AllTables**("_tablename_")|AllTables("OrderTable")|
|**AllTables**(_index_)|AllTables(0)|

## Properties

- [AllDatabaseDiagrams](Access.CodeData.AllDatabaseDiagrams.md)
- [AllFunctions](Access.CodeData.AllFunctions.md)
- [AllQueries](Access.CodeData.AllQueries.md)
- [AllStoredProcedures](Access.CodeData.AllStoredProcedures.md)
- [AllTables](Access.CodeData.AllTables.md)
- [AllViews](Access.CodeData.AllViews.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]