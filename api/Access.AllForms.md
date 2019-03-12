---
title: AllForms object (Access)
keywords: vbaac10.chm12683
f1_keywords:
- vbaac10.chm12683
ms.prod: access
api_name:
- Access.AllForms
ms.assetid: b90616b9-90fc-bb51-6bfa-b149dece0f1b
ms.date: 02/01/2019
localization_priority: Normal
---


# AllForms object (Access)

The **AllForms** collection contains an **[AccessObject](Access.AccessObject.md)** object for each form in the **[CurrentProject](Access.CurrentProject.md)** or **[CodeProject](Access.CodeProject.md)** object.


## Remarks

The **CurrentProject** and **CodeProject** object has an **AllForms** collection containing **AccessObject** objects that describe instances of all the forms in the database. For example, you can enumerate the **AllForms** collection in Visual Basic to set or return the values of properties of individual **AccessObject** objects in the collection.

You can refer to an individual **AccessObject** object in the **AllForms** collection either by referring to the object by name, or by referring to its index within the collection. If you want to refer to a specific object in the **AllForms** collection, it's better to refer to the form by name because a form's collection index may change.

The **AllForms** collection is indexed beginning with zero. If you refer to a form by its index, the first form is AllForms(0), the second form is AllForms(1), and so on.

> [!NOTE] 
> To list all open forms in the database, use the **[IsLoaded](Access.AccessObject.IsLoaded.md)** property of each **AccessObject** object in the **AllForms** collection. You can then use the **Name** property of each individual **AccessObject** object to return the name of a form.

You can't add or delete an **AccessObject** object from the **AllForms** collection.


## Example

The following example prints the name of each open **AccessObject** object in the **AllForms** collection.

```vb
Sub AllForms() 
    Dim obj As AccessObject, dbs As Object 
    Set dbs = Application.CurrentProject 
    ' Search for open AccessObject objects in AllForms collection. 
    For Each obj In dbs.AllForms 
        If obj.IsLoaded = True Then 
            ' Print name of obj. 
            Debug.Print obj.Name 
        End If 
    Next obj 
End Sub
```

<br/>

The following example shows how to prevent a user from opening a particular form directly from the navigation pane.

```vb
'Don't let this form be opened from the Navigator
If Not CurrentProject.AllForms(cFormUsage).IsLoaded Then
    MsgBox "This form cannot be opened from the navigation pane.", _
        vbInformation + vbOKOnly, "Invalid form usage"
    Cancel = True
    Exit Sub
End If
```


## Properties

- [Application](Access.AllForms.Application.md)
- [Count](Access.AllForms.Count.md)
- [Item](Access.AllForms.Item.md)
- [Parent](Access.AllForms.Parent.md)


## See also

- [Access Object Model Reference](overview/Access/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
