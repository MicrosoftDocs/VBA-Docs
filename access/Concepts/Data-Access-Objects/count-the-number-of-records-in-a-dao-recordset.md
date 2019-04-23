---
title: Count the number of records in a DAO Recordset
ms.prod: access
ms.assetid: ea524046-4d04-b595-1a45-13b399745f44
ms.date: 09/21/2018
localization_priority: Normal
---


# Count the number of records in a DAO Recordset

You may want to know the number of records in a **[Recordset](../../../api/overview/Access.md)** object. For example, you may want to create a form that shows how many records are in each of the tables in a database. Or you may want to change the appearance of a form or report based on the number of records it includes.

The **[RecordCount](../../../api/overview/Access.md)** property contains the number of records in a table-type **Recordset** or the total number of records accessed in a dynaset- or snapshot-type **Recordset**. A **Recordset** object with no records has a **RecordCount** property value of 0.

> [!NOTE] 
> The value of the **RecordCount** property equals the number of records that have actually been accessed. For example, when you first create a dynaset or snapshot, you have accessed (or visited) only one record. If you check the **RecordCount** property immediately after creating the dynaset or snapshot (assuming it has at least one record), the value is 1. To visit all the records, use the **[MoveLast](../../../api/overview/Access.md)** method immediately after opening the **Recordset**, and then use **[MoveFirst](../../../api/overview/Access.md)** to return to the first record. This is not done automatically because it may be slow, especially for large result sets.

When you open a table-type **Recordset** object, you effectively visit all of the records in the underlying table, and the value of the **RecordCount** property equals the number of records in the table as soon as the **Recordset** is opened.

Canceled transactions may make the value of the **RecordCount** property out-of-date in some multiuser situations. Compacting the database restores the table's record count to the correct value.

The following code example creates a snapshot-type **Recordset** object, and then determines the number of records in the **Recordset**.

```vb
Function FindRecordCount(strSQL As String) As Long 
 
Dim dbsNorthwind As DAO.Database 
Dim rstRecords As DAO.Recordset 
 
On Error GoTo ErrorHandler 
 
   Set dbsNorthwind = CurrentDb 
 
   Set rstRecords = dbsNorthwind.OpenRecordset(strSQL) 
 
   If rstRecords.EOF Then 
      FindRecordCount = 0 
   Else 
      rstRecords.MoveLast 
      FindRecordCount = rstRecords.RecordCount 
   End If 
 
   rstRecords.Close 
   dbsNorthwind.Close 
 
   Set rstRecords = Nothing 
   Set dbsNorthwind = Nothing 
 
Exit Function 
 
ErrorHandler: 
   MsgBox "Error #: " & Err.Number & vbCrLf & vbCrLf & Err.Description 
End Function
```

As your application deletes records in a dynaset-type **Recordset**, the value of the **RecordCount** property decreases. However, in a multiuser environment, records deleted by other users are not reflected in the value of the **RecordCount** property until the current record is positioned on a deleted record. At that time, the setting of the **RecordCount** property decreases by one. 

Using the **[Requery](../../../api/overview/Access.md)** method on a **Recordset**, followed by the **MoveLast** method, sets the **RecordCount** property to the current total number of records in the **Recordset**.

A snapshot-type **Recordset** object is static and the value of its **RecordCount** property does not change when you add or delete records in the snapshot's underlying table.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
