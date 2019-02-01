---
title: AllViews object (Access)
keywords: vbaac10.chm12690
f1_keywords:
- vbaac10.chm12690
ms.prod: access
api_name:
- Access.AllViews
ms.assetid: f56bee24-a972-fbdf-f74a-0ac83825e3bb
ms.date: 02/01/2019
localization_priority: Normal
---


# AllViews object (Access)

The **AllViews** collection contains an **[AccessObject](Access.AccessObject.md)** for each view in the **[CurrentData](Access.CurrentData.md)** or **[CodeData](Access.CodeData.md)** object.


## Remarks

The **CurrentData** or **CodeData** object has an **AllViews** collection containing **AccessObject** objects that describe instances of all views specified by **CurrentData** or **CodeData**. For example, you can enumerate the **AllViews** collection in Visual Basic to set or return the values of properties of individual **AccessObject** objects in the collection.

You can refer to an individual **AccessObject** object in the **AllViews** collection either by referring to the object by name, or by referring to its index within the collection. If you want to refer to a specific object in the **AllViews** collection, it's better to refer to the view by name because a view's collection index may change.

The **AllViews** collection is indexed beginning with zero. If you refer to a view by its index, the first view is AllViews(0), the second view is AllViews(1), and so on.

> [!NOTE] 
> To list all open views in the database, use the **[IsLoaded](Access.AccessObject.IsLoaded.md)** property of each **AccessObject** object in the **AllViews** collection. You can then use the **Name** property of each individual **AccessObject** object to return the name of a view.

You can't add or delete an **AccessObject** object from the **AllViews** collection.


## Example

The following example prints the name of each open **AccessObject** object in the **AllViews** collection.


```vb
Sub AllViews() 
    Dim obj As AccessObject, dbs As Object 
    Set dbs = Application.CurrentData 
    ' Search for open AccessObject objects in AllViews collection. 
    For Each obj In dbs.AllViews 
        If obj.IsLoaded = True Then 
            ' Print name of obj. 
            Debug.Print obj.Name 
        End If 
    Next obj 
End Sub
```


## Properties

- [Application](Access.AllViews.Application.md)
- [Count](Access.AllViews.Count.md)
- [Item](Access.AllViews.Item.md)
- [Parent](Access.AllViews.Parent.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]