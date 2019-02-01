---
title: AllQueries object (Access)
keywords: vbaac10.chm12689
f1_keywords:
- vbaac10.chm12689
ms.prod: access
api_name:
- Access.AllQueries
ms.assetid: 9b67f04c-2642-0dcc-2a64-8ca8fa7249b3
ms.date: 02/01/2019
localization_priority: Normal
---


# AllQueries object (Access)

The **AllQueries** collection contains an **[AccessObject](Access.AccessObject.md)** for each query in the **[CurrentData](Access.CurrentData.md)** or **[CodeData](Access.CodeData.md)** object.


## Remarks

The **CurrentData** or **CodeData** object has an **AllQueries** collection containing **AccessObject** objects that describe instances of all queries specified by **CurrentData** or **CodeData**. For example, you can enumerate the **AllQueries** collection in Visual Basic to set or return the values of properties of individual **AccessObject** objects in the collection.

You can refer to an individual **AccessObject** object in the **AllQueries** collection either by referring to the object by name, or by referring to its index within the collection. If you want to refer to a specific object in the **AllQueries** collection, it's better to refer to the query by name because a query's collection index may change.

The **AllQueries** collection is indexed beginning with zero. If you refer to a query by its index, the first query is AllQueries(0), the second query is AllQueries(1), and so on.

> [!NOTE] 
> To list all open queries in the database, use the **[IsLoaded](Access.AccessObject.IsLoaded.md)** property of each **AccessObject** object in the **AllQueries** collection. You can then use the **Name** property of each individual **AccessObject** object to return the name of a query.

You can't add or delete an **AccessObject** object from the **AllQueries** collection.


## Example

The following example prints the name of each open **AccessObject** object in the **AllQueries** collection.


```vb
Sub AllQueries() 
    Dim obj As AccessObject, dbs As Object 
    Set dbs = Application.CurrentData 
    ' Search for open AccessObject objects in AllQueries collection. 
    For Each obj In dbs.AllQueries 
        If obj.IsLoaded = True Then 
            ' Print name of obj. 
            Debug.Print obj.Name 
        End If 
    Next obj 
End Sub
```


## Properties

- [Application](Access.AllQueries.Application.md)
- [Count](Access.AllQueries.Count.md)
- [Item](Access.AllQueries.Item.md)
- [Parent](Access.AllQueries.Parent.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]