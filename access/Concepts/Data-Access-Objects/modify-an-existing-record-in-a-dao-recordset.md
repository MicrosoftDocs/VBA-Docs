---
title: Modify an existing record in a DAO Recordset
ms.prod: access
ms.assetid: e1fe83cc-db41-8c51-1809-e5ae059c0260
ms.date: 09/21/2018
localization_priority: Normal
---


# Modify an existing record in a DAO Recordset

You can modify existing records in a table-type or dynaset-type **[Recordset](../../../api/overview/Access.md)** object by using the **[Edit](../../../api/overview/Access.md)** and **[Update](../../../api/overview/Access.md)** methods.

To modify an existing record in a table-type or dynaset-type **Recordset** object:

1. Go to the record that you want to change.
    
2. Use the **Edit** method to prepare the current record for editing.
    
3. Make the necessary changes to the record.
    
4. Use the **Update** method to save the changes to the current record.
    
The following code example shows how to change the job titles for all sales representatives in a table called Employees.

```vb
Dim dbsNorthwind As DAO.Database 
Dim rstEmployees As DAO.Recordset 
 
   Set dbsNorthwind = CurrentDb 
   Set rstEmployees = dbsNorthwind.OpenRecordset("Employees") 
 
   rstEmployees.MoveFirst 
   Do Until rstEmployees.EOF 
      If rstEmployees!Title = "Sales Representative" Then 
         rstEmployees.Edit 
         rstEmployees!Title = "Account Executive" 
         rstEmployees.Update 
      End If 
      rstEmployees.MoveNext 
   Loop 

```


> [!NOTE] 
> If you do not use the **Edit** method before you try to change a value in the current record, a run-time error occurs. If you edit the current record and then move to another record or close the **Recordset** object without first using the **Update** method, your changes are lost without warning. For example, omitting the **Update** method from the preceding example results in no changes being made to the Employees table.
> 
> You can also terminate the **Edit** method and any pending transactions without saving changes by using the **[CancelUpdate](../../../api/overview/Access.md)** method. While you can terminate the **Edit** method just by moving off the current record, this is not practical when the current record is the first or last record in the **Recordset**, or when it is a new record. It is generally simpler to use the **CancelUpdate** method.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
