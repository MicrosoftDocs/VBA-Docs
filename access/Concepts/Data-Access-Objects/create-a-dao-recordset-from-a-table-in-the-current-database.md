---
title: Create a DAO Recordset from a table in the current database
ms.prod: access
ms.assetid: b0507965-e6af-cda4-9d50-fbeb98b4ab89
ms.date: 09/21/2018
localization_priority: Normal
---


# Create a DAO Recordset from a table in the current database

The following code example uses the **[OpenRecordset](../../../api/overview/Access.md)** method to create a table-type **[Recordset](../../../api/overview/Access.md)** object for a table in the current database.


```vb
Dim dbsNorthwind As DAO.Database 
Dim rstCustomers As DAO.Recordset 
 
Set dbsNorthwind = CurrentDb 
Set rstCustomers = dbsNorthwind.OpenRecordset("Customers") 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
