---
title: AllMacros object (Access)
keywords: vbaac10.chm12685
f1_keywords:
- vbaac10.chm12685
ms.prod: access
ms.assetid: a36ba978-f643-aca6-5efb-842723d17bbc
ms.date: 02/01/2019
localization_priority: Normal
---


# AllMacros object (Access)

The **AllMacros** collection contains an **[AccessObject](Access.AccessObject.md)** for each macro in the **[CurrentProject](Access.CurrentProject.md)** or **[CodeProject](Access.CodeProject.md)** object.


## Remarks

The **CurrentProject** or **CodeProject** object has an **AllMacros** collection containing **AccessObject** objects that describe instances of all the macros specified by **CurrentProject** or **CodeProject**. For example, you can enumerate the **AllMacros** collection in Visual Basic to set or return the values of properties of individual **AccessObject** objects in the collection.

You can refer to an individual **AccessObject** object in the **AllMacros** collection either by referring to the object by name, or by referring to its index within the collection. If you want to refer to a specific object in the **AllMacros** collection, it's better to refer to the macro by name because a macro's collection index may change.

The **AllMacros** collection is indexed beginning with zero. If you refer to a macro by its index, the first macro is AllMacros(0), the second macro is AllMacros(1), and so on.

> [!NOTE] 
> To list all open macros in the database, use the **[IsLoaded](Access.AccessObject.IsLoaded.md)** property of each **AccessObject** object in the **AllMacros** collection. You can then use the **Name** property of each individual **AccessObject** object to return the name of a macro.

You can't add or delete an **AccessObject** object from the **AllMacros** collection.


## Example

The following example prints the name of each open **AccessObject** object in the **AllMacros** collection.


```vb
Sub AllMacros() 
 Dim obj As AccessObject, dbs As Object 
 Set dbs = Application.CurrentProject 
 ' Search for open AccessObject objects in AllMacros collection. 
 For Each obj In dbs.AllMacros 
 If obj.IsLoaded = True Then 
 ' Print name of obj. 
 Debug.Print obj.Name 
 End If 
 Next obj 
End Sub
```


## Properties

- [Application](Access.AllMacros.Application.md)
- [Count](Access.AllMacros.Count.md)
- [Item](Access.AllMacros.Item.md)
- [Parent](Access.AllMacros.Parent.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]