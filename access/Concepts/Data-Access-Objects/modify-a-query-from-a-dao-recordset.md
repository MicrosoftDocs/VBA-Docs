---
title: Modify a query from a DAO Recordset
ms.prod: access
ms.assetid: b5679ca8-9bcd-2d28-15af-2640db727dd4
ms.date: 09/21/2018
localization_priority: Normal
---


# Modify a query from a DAO Recordset

You can use the **[Requery](../../../api/overview/Access.md)** method on a dynaset-type or snapshot-type **[Recordset](../../../api/overview/Access.md)** object when you want to run the underlying query again after changing a parameter. This is more convenient than opening a new **Recordset**, and it runs faster.

The following code example creates a **Recordset** object and passes it to a function that uses the **[CopyQueryDef](../../../api/overview/Access.md)** method to extract the equivalent SQL string. It then prompts the user to add an additional constraint clause to the query. The code uses the **Requery** method to run the modified query.



```vb
Sub AddQuery() 
 
Dim dbsNorthwind As DAO.Database 
Dim qdfSalesReps As DAO.QueryDef 
Dim rstSalesReps As DAO.Recordset 
 
On Error GoTo ErrorHandler 
 
   Set dbsNorthwind = CurrentDb 
 
   Set qdfSalesReps = dbsNorthwind.CreateQueryDef("SalesRepQuery") 
   qdfSalesReps.SQL = "SELECT * FROM Employees WHERE Title = " & _ 
                      "'Sales Representative'" 
 
   Set rstSalesReps = qdfSalesReps.OpenRecordset() 
 
   ' Call the function to add a constraint. 
   AddQueryFilter rstSalesReps 
 
   ' Return database to original. 
   dbsNorthwind.QueryDefs.Delete "SalesRepQuery" 
 
   rstSalesReps.Close 
   qdfSalesReps.Close 
   dbsNorthwind.Close 
 
   Set rstSalesReps = Nothing 
   Set qdfSalesReps = Nothing 
   Set dbsNorthwind = Nothing 
 
   Exit Sub 
 
ErrorHandler: 
   MsgBox "Error #: " & Err.Number & vbCrLf & vbCrLf & Err.Description 
End Sub 
 
Sub AddQueryFilter(rstData As Recordset) 
 
Dim qdfData As DAO.QueryDef 
Dim strNewFilter As String 
Dim strRightSQL As String 
 
On Error GoTo ErrorHandler 
 
   Set qdfData = rstData.CopyQueryDef 
 
   ' Try "LastName LIKE 'D*'". 
   strNewFilter = InputBox("Enter new criteria") 
 
   strRightSQL = Right(qdfData.SQL, 1) 
 
   ' Strip characters from the end of the query, 
   ' as needed. 
   Do While strRightSQL = " " Or strRightSQL = ";" Or _ 
                          strRightSQL = vbCR Or strRightSQL = vbLF 
      qdfData.SQL = Left(qdfData.SQL, Len(qdfData.SQL) - 1) 
      strRightSQL = Right(qdfData.SQL, 1) 
   Loop 
 
   qdfData.SQL = qdfData.SQL & " AND " & strNewFilter 
   rstData.Requery qdfData         'Requery the Recordset. 
   rstData.MoveLast               'Populate the Recordset. 
 
   ' "Lastname LIKE 'D*'" should return 2 records. 
   MsgBox "Number of records found:  " & rstData.RecordCount & "." 
 
   qdfData.Close 
   Set qdfData = Nothing 
 
   Exit Sub 
 
ErrorHandler: 
   MsgBox "Error #: " & Err.Number & vbCrLf & vbCrLf & Err.Description 
End Sub
```

> [!NOTE] 
> To use the **Requery** method, the **[Restartable](../../../api/overview/Access.md)** property of the **Recordset** object must be set to **True**. The **Restartable** property is always set to **True** when the **Recordset** is created from a query other than a crosstab query against tables in an Access database. You cannot restart SQL pass-through queries. You may or may not be able to restart queries against linked tables in another database format. To determine whether a **Recordset** object can rerun its query, check the **Restartable** property.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]