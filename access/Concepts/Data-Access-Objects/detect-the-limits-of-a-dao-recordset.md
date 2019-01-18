---
title: Detect the limits of a DAO Recordset
ms.prod: access
ms.assetid: f4be9ea8-25af-1c5c-4cd7-43d57e5d4d8b
ms.date: 09/21/2018
localization_priority: Normal
---


# Detect the limits of a DAO Recordset

In a **[Recordset](../../../api/overview/Access.md)** object, if you try to move beyond the beginning or ending record, a run-time error occurs. For example, if you try to use the **[MoveNext](../../../api/overview/Access.md)** method when you are already at the last record of the **Recordset**, a trappable error occurs. For this reason, it is helpful to know the limits of the **Recordset** object.

The **[BOF](../../../api/overview/Access.md)** property indicates whether the current position is at the beginning of the **Recordset**. If **BOF** is **True**, the current position is before the first record in the **Recordset**. The **BOF** property is also **True** if there are no records in the **Recordset** when it is opened. 

Similarly, the **[EOF](../../../api/overview/Access.md)** property is **True** if the current position is after the last record in the **Recordset**, or if there are no records.

The following code example shows how to use the **BOF** and **EOF** properties to detect the beginning and end of a **Recordset** object. This code fragment creates a table-type **Recordset** based on the Orders table from the current database. It moves through the records, first from the beginning of the **Recordset** to the end, and then from the end of the **Recordset** to the beginning.

```vb
Dim dbsNorthwind As DAO.Database 
Dim rstOrders As DAO.Recordset 
 
   Set dbsNorthwind = CurrentDb 
   Set rstOrders = dbsNorthwind.OpenRecordset("Orders") 
 
   ' Do until ending of file. 
   Do Until rstOrders.EOF 
      ' 
      ' Manipulate the data. 
      ' 
      rstOrders.MoveNext            ' Move to the next record. 
   Loop 
 
   rstOrders.MoveLast               ' Move to the last record. 
 
   ' Do until beginning of file. 
   Do Until rstOrders.BOF 
      ' 
      ' Manipulate the data. 
      ' 
      rstOrders.MovePrevious        ' Move to the previous record. 
   Loop 

```

Be aware that there is no current record immediately following the first loop. The **BOF** and **EOF** properties both have the following characteristics.

- If the **Recordset** contains no records when you open it, both **BOF** and **EOF** are **True**.
    
- When **BOF** or **EOF** is **True**, the property remains **True** until you move to an existing record, at which time the value of **BOF** or **EOF** becomes **False**.
    
- When **BOF** or **EOF** is **False**, and the only record in a **Recordset** is deleted, the property remains **False** until you try to move to another record, at which time both **BOF** and **EOF** become **True**.
    
- At the moment you create or open a **Recordset** that contains at least one record, the first record is the current record, and both **BOF** and **EOF** are **False**.
    
- If the first record is the current record when you use the **MovePrevious** method, **BOF** is set to **True**. If you use **MovePrevious** while **BOF** is **True**, a run-time error occurs. When this happens, **BOF** remains **True** and there is no current record.
    
- Similarly, moving past the last record in the **Recordset** changes the value of the **EOF** property to **True**. If you use the **MoveNext** method while **EOF** is **True**, a run-time error occurs. When this happens, **EOF** remains **True** and there is no current record.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]