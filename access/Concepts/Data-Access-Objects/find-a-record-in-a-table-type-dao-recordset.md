---
title: Find a record in a table-type DAO Recordset
ms.prod: access
ms.assetid: b17f14db-9b3e-7f12-9fc8-f56c6dcbad09
ms.date: 09/21/2018
localization_priority: Normal
---


# Find a record in a table-type DAO Recordset

You use the **[Seek](../../../api/overview/Access.md)** method to locate a record in a table-type **[Recordset](../../../api/overview/Access.md)** object.

When you use the **Seek** method to locate a record, the Access database engine uses the table's current index, as defined by the **[Index](../../../api/overview/Access.md)** property.

> [!NOTE] 
> If you use the **Seek** method on a table-type **Recordset** object without first setting the current index, a run-time error occurs.

The following example opens a table-type **Recordset** object called Employees, and uses the Seek method to locate the record containing a value of **lngEmpID** in the EmployeeID field. It returns the hire date for the specified employee.



```vb
Function GetHireDate(lngEmpID As Long) As Variant 
 
Dim dbsNorthwind As DAO.Database 
Dim rstEmployees As DAO.Recordset 
 
On Error GoTo ErrorHandler 
 
   Set dbsNorthwind = CurrentDB 
   Set rstEmployees = dbsNorthwind.OpenRecordset("Employees") 
 
   ' The index name for Employee ID. 
   rstEmployees.Index = "PrimaryKey" 
   rstEmployees.Seek "=", lngEmpID 
 
   If rstEmployees.NoMatch Then 
      GetHireDate = Null 
   Else 
      GetHireDate = rstEmployees!HireDate 
   End If 
 
   rstEmployees.Close 
   dbsNorthwind.Close 
 
   Set rstEmployees = Nothing 
   Set dbsNorthwind = Nothing 
 
Exit Function 
 
ErrorHandler: 
   MsgBox "Error #: " & Err.Number & vbCrLf & vbCrLf & Err.Description 
End Function
```

The **Seek** method always starts searching for records at the beginning of the **Recordset** object. If you use the **Seek** method with the same arguments more than once on the same **Recordset**, it finds the same record.

You can use the **[NoMatch](../../../api/overview/Access.md)** property on the **Recordset** object to test whether a record matching the search criteria was found. If the record matching the criteria was found, the **NoMatch** property will be **False**; otherwise, it will be **True**.

The following code example shows how you can create a function that uses the **Seek** method to locate a record by using a multiple-field index.

```vb
Function GetFirstPrice(lngOrderID As Long, lngProductID As Long) As Variant 
 
Dim dbsNorthwind As DAO.Database 
Dim rstOrderDetail As DAO.Recordset 
 
On Error GoTo ErrorHandler 
 
   Set dbsNorthwind = CurrentDb 
   Set rstOrderDetail = dbsNorthwind.OpenRecordset("Order Details") 
 
   rstOrderDetail.Index = "PrimaryKey" 
   rstOrderDetail.Seek "=", lngOrderID, lngProductID 
 
   If rstOrderDetail.NoMatch Then 
      GetFirstPrice = Null 
   Else 
      GetFirstPrice = rstOrderDetail!UnitPrice 
   End If 
 
   rstOrderDetail.Close 
   dbsNorthwind.Close 
 
   Set rstOrderDetail = Nothing 
   Set dbsNorthwind = Nothing 
 
Exit Function 
 
ErrorHandler: 
   MsgBox "Error #: " & Err.Number & vbCrLf & vbCrLf & Err.Description 
End Function
```

In this example, the table's primary key consists of two fields: OrderID and ProductID. When you call the GetFirstPrice function with a valid (existing) combination of OrderID and ProductID field values, the function returns the unit price from the found record. If it cannot find the combination of field values you want in the table, the function returns the **Null** value.

If the current index is a multiple-field index, trailing key values can be omitted and are treated as **Null** values. That is, you can leave off any number of key values from the end of a **Seek** method's _key_ argument, but not from the beginning or the middle. However, if you do not specify all values in the index, you can use only the ">" or "<" comparison string with the **Seek** method.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
