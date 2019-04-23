---
title: AllReports object (Access)
keywords: vbaac10.chm12684
f1_keywords:
- vbaac10.chm12684
ms.prod: access
api_name:
- Access.AllReports
ms.assetid: 5846cf60-41b4-e9f8-ea27-b9400a6d3861
ms.date: 02/01/2019
localization_priority: Normal
---


# AllReports object (Access)

The **AllReports** collection contains an **[AccessObject](Access.AccessObject.md)** for each report in the **[CurrentProject](Access.CurrentProject.md)** or **[CodeProject](Access.CodeProject.md)** object.


## Remarks

The **CurrentProject** or **CodeProject** object has an **AllReports** collection containing **AccessObject** objects that describe instances of all the reports in the database. For example, you can enumerate the **AllReports** collection in Visual Basic to set or return the values of properties of individual **AccessObject** objects in the collection.

You can refer to an individual **AccessObject** object in the **AllReports** collection either by referring to the item by name, or by referring to its index within the collection. If you want to refer to a specific report in the **AllReports** collection, it's better to refer to the report by name because the index may change.

The **AllReports** collection is indexed beginning with zero. If you refer to a report by its index, the first report is AllReports(0), the second report is AllReports(1), and so on.

> [!NOTE] 
> To list all open reports in the database, use the **[IsLoaded](Access.AccessObject.IsLoaded.md)** property of each **AccessObject** object in the **AllReports** collection. You can then use the **Name** property of each individual **AccessObject** object to return the name of a report.

You can't add or delete an **AccessObject** object from the **AllReports** collection.


## Example

The following example prints the name of each open **AccessObject** object in the **AllReports** collection.


```vb
Sub AllReports() 
 Dim obj As AccessObject, dbs As Object 
 Set dbs = Application.CurrentProject 
 ' Search for open AccessObject objects in AllReports collection. 
 For Each obj In dbs.AllReports 
 If obj.IsLoaded = True Then 
 ' Print name of obj. 
 Debug.Print obj.Name 
 End If 
 Next obj 
End Sub
```


## Properties

- [Application](Access.AllReports.Application.md)
- [Count](Access.AllReports.Count.md)
- [Item](Access.AllReports.Item.md)
- [Parent](Access.AllReports.Parent.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]