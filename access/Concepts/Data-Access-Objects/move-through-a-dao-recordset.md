---
title: Move through a DAO Recordset
ms.prod: access
ms.assetid: 7d788b60-c6e8-dea7-68fe-01b893fc3374
ms.date: 09/21/2018
localization_priority: Normal
---


# Move through a DAO Recordset

A **[Recordset](../../../api/overview/Access.md)** object usually has a current position, most often at a record. When you refer to the fields in a **Recordset**, you obtain values from the record at the current position, which is known as the current record. However, the current position can also be immediately before the first record in a **Recordset** or immediately after the last record. In certain circumstances, the current position is undefined.

You can use the following **Move** methods to loop through the records in a **Recordset**:

- The **[MoveFirst](../../../api/overview/Access.md)** method moves to the first record.
    
- The **[MoveLast](../../../api/overview/Access.md)** method moves to the last record.
    
- The **[MoveNext](../../../api/overview/Access.md)** method moves to the next record.
    
- The **[MovePrevious](../../../api/overview/Access.md)** method moves to the previous record.
    
- The **[Move](../../../api/overview/Access.md)** method moves forward or backward the number of records you specify in its syntax.
    
You can use each of these methods on table-type, dynaset-type, and snapshot-type **Recordset** objects. On a forward-only-type **Recordset** object, you can use only the **MoveNext** and **Move** methods. If you use the **Move** method on a forward-only-type **Recordset**, the argument specifying the number of rows to move must be a positive integer.

The following code example opens a **Recordset** object on the Employees table containing all of the records that have a **Null** value in the ReportsTo field. The function then updates the records to indicate that these employees are temporary employees. For each record in the **Recordset**, the example changes the Title and Notes fields, and saves the changes with the **[Update](../../../api/overview/Access.md)** method. It uses the **MoveNext** method to move to the next record.

```vb
Sub UpdateEmployees() 
 
Dim dbsNorthwind As DAO.Database 
Dim rstEmployees As DAO.Recordset 
Dim strSQL As String 
Dim intI As Integer 
 
On Error GoTo ErrorHandler 
 
   Set dbsNorthwind = CurrentDb 
 
   ' Open a recordset on all records from the Employees table that have 
   ' a Null value in the ReportsTo field. 
   strSQL = "SELECT * FROM Employees WHERE ReportsTo IS NULL" 
   Set rstEmployees = dbsNorthwind.OpenRecordset(strSQL, dbOpenDynaset) 
 
   ' If the recordset is empty, exit. 
   If rstEmployees.EOF Then Exit Sub 
 
   intI = 1 
   With rstEmployees 
      Do Until .EOF 
         .Edit 
         ![ReportsTo] = 5 
         ![Title] = "Temporary" 
         ![Notes] = rstEmployees![Notes] & "Temp #" & intI 
         .Update 
         .MoveNext 
         intI = intI + 1 
      Loop 
   End With 
 
   RstEmployees.Close 
   dbsNorthwind.Close 
 
   Set rstEmployees = Nothing 
   Set dbsNorthwind = Nothing 
 
   Exit Sub 
 
ErrorHandler: 
   MsgBox "Error #: " & Err.Number & vbCrLf & vbCrLf & Err.Description 
End Sub
```

> [!NOTE] 
> The previous example is provided only for the purposes of illustrating the **Update** and **MoveNext** methods. For optimal performance, it is recommended that you perform this bulk operation with a SQL UPDATE query.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
