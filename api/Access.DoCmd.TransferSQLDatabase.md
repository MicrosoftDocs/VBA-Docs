---
title: DoCmd.TransferSQLDatabase method (Access)
keywords: vbaac10.chm5085
f1_keywords:
- vbaac10.chm5085
ms.prod: access
api_name:
- Access.DoCmd.TransferSQLDatabase
ms.assetid: d6a88496-9137-b190-8357-316fd580a036
ms.date: 03/07/2019
localization_priority: Normal
---


# DoCmd.TransferSQLDatabase method (Access)

Transfers the entire specified Microsoft SQL Server database to another SQL Server database.


## Syntax

_expression_.**TransferSQLDatabase** (_Server_, _Database_, _UseTrustedConnection_, _Login_, _Password_, _TransferCopyData_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Server_|Required|**Variant**|The name of the SQL Server to which the database will be transferred.|
| _Database_|Required|**Variant**|The name of the new database on the specified server.|
| _UseTrustedConnection_|Optional|**Variant**|**True** if the current connection is using a login with system administrator privileges. If this argument is not **True**, you must specify a login and password in the _Login_ and _Password_ arguments.|
| _Login_|Optional|**Variant**|The name of a login on the destination server with system administrator privileges. If  _UseTrustedConnection_ is **True**, this argument is ignored.|
| _Password_|Optional|**Variant**|The password for the login specified in _Login_. If _UseTrustedConnection_ is **True**, this argument is ignored.|
| _TransferCopyData_|Optional|**Variant**|**True** if all data in the database is transferred to the destination database. If this argument is not **True**, only the database schema will be transferred.|

## Remarks

The following conditions must be met or else an error occurs:

- The current and destination servers are SQL Server version 7.0 or later.
    
- The user has system administrator login rights on the destination server.
    
- The destination database doesn't already exist on the destination server.
    

## Example

This example transfers the current SQL Server database to a new SQL Server database called Inventory on the server MainOffice. (It is assumed that the user has system administrator privileges on MainOffice.) The data is copied along with the database schema.

```vb
DoCmd.TransferCompleteSQLDatabase _ 
 Server:="MainOffice", _ 
 Database:="Inventory", _ 
 UseTrustedConnection:=True, _ 
 TransferCopyData:=False 

```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]