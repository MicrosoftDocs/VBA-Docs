---
title: AllDatabaseDiagrams object (Access)
keywords: vbaac10.chm12692
f1_keywords:
- vbaac10.chm12692
ms.prod: access
api_name:
- Access.AllDatabaseDiagrams
ms.assetid: 417427aa-1783-29da-30c9-66a7032a0088
ms.date: 02/01/2019
localization_priority: Normal
---


# AllDatabaseDiagrams object (Access)

The **AllDatabaseDiagrams** collection contains an **[AccessObject](Access.AccessObject.md)** for each database diagram in the **[CurrentData](Access.CurrentData.md)** or **[CodeData](Access.CodeData.md)** object.


## Remarks

The **CurrentData** or **CodeData** object has an **AllDatabaseDiagrams** collection containing **AccessObject** objects that describe instances of all database diagrams specified by **CurrentData** or **CodeData**. For example, you can enumerate the **AllDatabaseDiagrams** collection in Visual Basic to set or return the values of properties of individual **AccessObject** objects in the collection.

You can refer to an individual **AccessObject** object in the **AllDatabaseDiagrams** collection either by referring to the object by name, or by referring to its index within the collection. If you want to refer to a specific object in the **AllDatabaseDiagrams** collection, it's better to refer to the database diagram by name because a database diagram's collection index may change.

The **AllDatabaseDiagrams** collection is indexed beginning with zero. If you refer to a database diagram by its index, the first database diagram is AllDatabaseDiagrams(0), the second database diagram is AllDatabaseDiagrams(1), and so on.

> [!NOTE] 
> You can't add or delete an **AccessObject** object from the **AllDatabaseDiagrams** collection.


## Example

The following example prints the name of each open **AccessObject** object in the **AllDatabaseDiagrams** collection.


```vb
Sub AllDatabaseDiagrams() 
    Dim obj As AccessObject, dbs As Object 
    Set dbs = Application.CurrentData 
    ' Search for open AccessObject objects in 
    ' AllDatabaseDiagrams collection. 
    For Each obj In dbs.AllDatabaseDiagrams 
        If obj.IsLoaded = True Then 
            ' Print name of obj. 
            Debug.Print obj.Name 
        End If 
    Next obj 
End Sub 
```


## Properties

- [Application](Access.AllDatabaseDiagrams.Application.md)
- [Count](Access.AllDatabaseDiagrams.Count.md)
- [Item](Access.AllDatabaseDiagrams.Item.md)
- [Parent](Access.AllDatabaseDiagrams.Parent.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]