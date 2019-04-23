---
title: Report.RecordLocks property (Access)
keywords: vbaac10.chm13696
f1_keywords:
- vbaac10.chm13696
ms.prod: access
api_name:
- Access.Report.RecordLocks
ms.assetid: 21f8d145-e417-a7a1-e697-b1e07434c760
ms.date: 03/20/2019
localization_priority: Normal
---


# Report.RecordLocks property (Access)

You can use the **RecordLocks** property to determine how records are locked and what happens when two users try to edit the same record at the same time. Read/write.


## Syntax

_expression_.**RecordLocks**

_expression_ A variable that represents a **[Report](Access.Report.md)** object.


## Remarks

When you edit a record, Microsoft Access can automatically lock that record to prevent other users from changing it before you are finished. For reports, the **RecordLocks** property specifies whether records in the underlying table or query are locked while a report is previewed or printed.

The **RecordLocks** property only applies to forms, reports, or queries in an Access database.

The **RecordLocks** property uses the following settings.

|Setting|Visual Basic|Description|
|:-----|:-----|:-----|
|No Locks|0|(Default) In reports, records aren't locked while the report is previewed or printed. In queries, records aren't locked while the query is run. This is also called "optimistic" locking.|
|All Records|1|All records in the underlying table or query are locked while the report is previewed or printed or the query is run. Although users can read the records, no one can edit, add, or delete any records until the report or query is closed.|
|Edited Record|2| Applies only to forms and queries. A page of records is locked as soon as any user starts editing any field in the record and stays locked until the user moves to another record. Consequently, a record can be edited by only one user at a time. This is also called "pessimistic" locking.|

> [!NOTE] 
> Changing the **RecordLocks** property of an open form or report causes an automatic recreation of the recordset.

You can use the No Locks setting for forms if only one person uses the underlying tables or queries or makes all the changes to the data.

In a multiuser database, you can use the No Locks setting if you want to use optimistic locking and warn users attempting to edit the same record on a form. You can use the Edited Record setting if you want to prevent two or more users from editing data at the same time.

You can use the All Records setting when you need to ensure that no changes are made to data after you start to preview or print a report or run an append, delete, make-table, or update-query.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]