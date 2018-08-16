---
title: Create a DAO Recordset From a Table In the Current Database
ms.prod: access
ms.assetid: b0507965-e6af-cda4-9d50-fbeb98b4ab89
ms.date: 06/08/2017
---


# Create a DAO Recordset From a Table In the Current Database

The following code example uses the  **[OpenRecordset](../../../api/overview/Access.md)** method to create a table-type **[Recordset](../../../api/overview/Access.md)** object for a table in the current database.


```vb
Dim dbsNorthwind As DAO.Database 
Dim rstCustomers As DAO.Recordset 
 
Set dbsNorthwind = CurrentDb 
Set rstCustomers = dbsNorthwind.OpenRecordset("Customers") 

```


