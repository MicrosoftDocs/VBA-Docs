---
title: AllStoredProcedures object (Access)
keywords: vbaac10.chm12691
f1_keywords:
- vbaac10.chm12691
ms.prod: access
api_name:
- Access.AllStoredProcedures
ms.assetid: 896f4c2c-273c-2849-0f06-d75fa515c44a
ms.date: 02/01/2019
localization_priority: Normal
---


# AllStoredProcedures object (Access)

The **AllStoredProcedures** collection contains an **[AccessObject](Access.AccessObject.md)** for each stored procedure in the **[CurrentData](Access.CurrentData.md)** or **[CodeData](Access.CodeData.md)** object.


## Remarks

The **CurrentData** or **CodeData** object has an **AllStoredProcedures** collection containing **AccessObject** objects that describe instances of all stored procedures specified by **CurrentData** or **CodeData**. For example, you can enumerate the **AllStoredProcedures** collection in Visual Basic to set or return the values of properties of individual **AccessObject** objects in the collection.

You can refer to an individual **AccessObject** object in the **AllStoredProcedures** collection either by referring to the object by name, or by referring to its index within the collection. If you want to refer to a specific object in the **AllStoredProcedures** collection, it's better to refer to the stored procedures by name because a stored procedure's collection index may change.

The **AllStoredProcedures** collection is indexed beginning with zero. If you refer to a stored procedure by its index, the first stored procedure is AllStoredProcedures(0), the second stored procedure is AllStoredProcedures(1), and so on.

> [!NOTE] 
> To list all open stored procedures in the database, use the **[IsLoaded](Access.AccessObject.IsLoaded.md)** property of each **AccessObject** object in the **AllStoredProcedures** collection. You can then use the **Name** property of each individual **AccessObject** object to return the name of a stored procedure.

You can't add or delete an **AccessObject** object from the **AllStoredProcedures** collection. 


## Example

The following example prints the name of each open **AccessObject** object in the **AllProcedures** collection.


```vb
Sub AllStoredProcedures() 
    Dim obj As AccessObject, dbs As Object 
    Set dbs = Application.CurrentData 
    ' Search for open AccessObject objects in 
    ' AllStoredProcedures collection. 
    For Each obj In dbs.AllStoredProcedures 
        If obj.IsLoaded = True Then 
            ' Print name of obj. 
            Debug.Print obj.Name 
        End If 
    Next obj 
End Sub
```


## Properties

- [Application](Access.AllStoredProcedures.Application.md)
- [Count](Access.AllStoredProcedures.Count.md)
- [Item](Access.AllStoredProcedures.Item.md)
- [Parent](Access.AllStoredProcedures.Parent.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]