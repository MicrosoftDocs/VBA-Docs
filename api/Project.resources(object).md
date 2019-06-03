---
title: Resources object (Project)
ms.prod: project-server
ms.assetid: 84f8357a-358b-f2ae-e164-65c0c5abd383
ms.date: 06/08/2017
localization_priority: Normal
---


# Resources object (Project)

Contains a collection of  **[Resource](Project.Resource.md)** objects.


## Example

 **Using the Resources Collection**

Use  **Resources** (_index_), where _index_ is the resource index number or resource name, to return a single **Resource** object. The following example lists the names of all resources in the active project.




```vb
Dim R As Long, Names As String 

 

For R = 1 To ActiveProject.Resources.Count 

 Names = ActiveProject.Resources(R).Name & ", " & Names 

Next R 

 

Names = Left$(Names, Len(Names) - Len(ListSeparator & " ")) 

MsgBox Names
```

 **Using the Resources Collection**

Use the  **[Resources](./Project.Project.Resources.md)** property to return a **Resources** collection. The following example generates the same list as the previous example, but does so by setting an object reference to `ActiveProject.Resources` , and then using `R` where `ActiveProject.Resources` is used.




```vb
Dim R As Resources, Temp As Long, Names As String 

 

Set R = ActiveProject.Resources 

 

For Temp = 1 To R.Count 

 Names = R(Temp).Name & ", " & Names 

Next Temp 

 

Names = Left$(Names, Len(Names) - Len(ListSeparator & " ")) 

MsgBox Names
```

Use the  **[Add](./Project.Resources.Add.md)** method to add a **Resource** object to the **Resources** collection. The following example adds a new resource named Matilda to the active project.




```vb
ActiveProject.Resources.Add "Matilda"
```


## Methods



|Name|
|:-----|
|[Add](./Project.Resources.Add.md)|

## Properties



|Name|
|:-----|
|[Application](./Project.Resources.Application.md)|
|[Count](./Project.Resources.Count.md)|
|[Item](./Project.Resources.Item.md)|
|[Parent](./Project.Resources.Parent.md)|
|[UniqueID](./Project.Resources.UniqueID.md)|

## See also


[Project Object Model](../project/Concepts/project-object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]