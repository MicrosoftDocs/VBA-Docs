---
title: Read from and write to a field in a DAO Recordset
ms.prod: access
ms.assetid: 4fe0c334-9c44-773c-7aed-182b042213a7
ms.date: 09/21/2018
localization_priority: Normal
---


# Read from and write to a field in a DAO Recordset

When you read or write data to a field, you are actually reading or setting the DAO **[Value](../../../api/overview/Access.md)** property of a **[Field](../../../api/overview/Access.md)** object. The DAO **Value** property is the default property of a **Field** object. Therefore, you can set the DAO **Value** property of the LastName field in the rstEmployees **[Recordset](../../../api/overview/Access.md)** in any of the following ways.


```vb
rstEmployees!LastName.Value = strName 
rstEmployees!LastName = strName 
rstEmployees![LastName] = strName 

```


The tables underlying a **Recordset** object may not permit you to modify data, even though the **Recordset** is of type dynaset or table, which are usually updatable. Check the **[Updatable](../../../api/overview/Access.md)** property of the **Recordset** to determine whether its data can be changed. If the property is **True**, the **Recordset** object can be updated.

Individual fields within an updatable **Recordset** object may not be updatable, and trying to write to these fields generates a run-time error. To determine whether a given field is updatable, check the **[DataUpdatable](../../../api/overview/Access.md)** property of the corresponding **Field** object in the **[Fields](../../../api/overview/Access.md)** collection of the **Recordset**. The following example returns **True** if all fields in the dynaset created by strQuery are updatable and returns **False** otherwise.



```vb
Function RecordsetUpdatable(strSQL As String) As Boolean 
 
Dim dbsNorthwind As DAO.Database 
Dim rstDynaset As DAO.Recordset 
Dim intPosition As Integer 
 
On Error GoTo ErrorHandler 
 
   ' Initialize the function's return value to True. 
   RecordsetUpdatable = True 
 
   Set dbsNorthwind = CurrentDb 
   Set rstDynaset = dbsNorthwind.OpenRecordset(strSQL, dbOpenDynaset) 
 
   ' If the entire dynaset isn't updatable, return False. 
   If rstDynaset.Updatable = False Then 
      RecordsetUpdatable = False 
   Else 
      ' If the dynaset is updatable, check if all fields in the 
      ' dynaset are updatable. If one of the fields isn't updatable, 
      ' return False. 
      For intPosition = 0 To rstDynaset.Fields.Count - 1 
         If rstDynaset.Fields(intPosition).DataUpdatable = False Then 
            RecordsetUpdatable = False 
            Exit For 
         End If 
      Next intPosition 
   End If 
 
   rstDynaset.Close 
   dbsNorthwind.Close 
 
   Set rstDynaset = Nothing 
   Set dbsNorthwind = Nothing 
 
Exit Sub 
 
ErrorHandler: 
   MsgBox "Error #: " & Err.Number & vbCrLf & vbCrLf & Err.Description 
End Function
```

Any single field can impose a number of criteria on data in that field when records are added or updated. These criteria are defined by a handful of properties. The DAO **[AllowZeroLength](../../../api/overview/Access.md)** property on a Text or Memo field indicates whether or not the field will accept a zero-length string (""). The DAO **[Required](../../../api/overview/Access.md)** property indicates whether or not some value must be entered in the field, or if it instead can accept a **Null** value. For a **Field** object on a **Recordset**, these properties are read-only; their state is determined by the underlying table.

Validation is the process of determining whether data entered into a field's DAO **Value** property is within an acceptable range. A **Field** object on a **Recordset** may have the DAO **[ValidationRule](../../../api/overview/Access.md)** and **[ValidationText](../../../api/overview/Access.md)** properties set. The DAO **ValidationRule** property is simply a criteria expression, similar to the criteria of an SQL WHERE clause, without the WHERE keyword. The DAO **ValidationText** property is a string that Access displays in an error message if you try to enter data in the field that is outside the limits of the DAO **ValidationRule** property. If you are using DAO in your code, then you can use the DAO **ValidationText** for a message that you want to display to the user.

> [!NOTE] 
> The DAO **ValidationRule** and **ValidationText** properties also exist at the **Recordset** level. These are read-only properties, reflecting the table-level validation scheme established on the table from which the current record is retrieved.

A **Field** object on a **Recordset** also features the **[ValidateOnSet](../../../api/overview/Access.md)** property. When the **ValidateOnSet** property is set to **True**, Access checks validation as soon as the field's DAO **Value** property is set. When it is set to **False** (the default), Access checks validation only when the completed record is updated. 

For example, if you are adding data to a record that contains a large Memo or OLE Object field and that has the DAO **ValidationRule** property set, you should determine whether the new data violates the validation rule before trying to write the data. To do so, set the **ValidateOnSet** property to **True**. If you wait to check validation until the entire record is written to disk, you may waste time trying to write an invalid record to disk.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
