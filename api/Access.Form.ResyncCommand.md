---
title: Form.ResyncCommand property (Access)
keywords: vbaac10.chm13486
f1_keywords:
- vbaac10.chm13486
ms.prod: access
api_name:
- Access.Form.ResyncCommand
ms.assetid: 0df53ea9-5771-0ccd-07ef-f33ad1082a61
ms.date: 03/15/2019
localization_priority: Normal
---


# Form.ResyncCommand property (Access)

You can use the **ResyncCommand** property to specify or determine the SQL statement or stored procedure that will be used in an updateable snapshot of a table. Read/write **String**.


## Syntax

_expression_.**ResyncCommand**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

The **ResyncCommand** property is a string expression representing a SQL statement or stored procedure that is parameterized by the key columns from the Unique Table in the output cursor.

The parameters must match in number and ordering to the set of key columns for the table identified by the **[UniqueTable](Access.Form.UniqueTable.md)** property. The purpose of the **ResyncCommand** property is to pull in the "fixed up" values of a row in a recordset after an update has been made, including an update to a join column.

For data access pages and forms based on views or non-parameterized SQL queries containing a join, if the **ResyncCommand** property is **Null**, Microsoft Access determines an appropriate query to use for the resink operation. 

For data access pages and forms based on stored procedures or parameterized SQL statements, Access cannot determine an appropriate resync query at run time, so the user must supply the **ResyncCommand** string to get the correct row fixup behavior. If the **ResyncCommand** property is empty and Access cannot determine an appropriate query to use, the default ADO resync operation (to display the current values) occurs after an update or insert.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]