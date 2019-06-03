---
title: Subprojects object (Project)
ms.prod: project-server
ms.assetid: 15688529-6d9c-6429-0d22-a5a16c033dcc
ms.date: 11/09/2018
localization_priority: Normal
---


# Subprojects object (Project)

Contains a collection of **[Subproject](Project.Subproject.md)** objects.

## Properties

|Name|
|:-----|
|[Application](./Project.Subprojects.Application.md)|
|[Count](./Project.Subprojects.Count.md)|
|[Item](./Project.Subprojects.Item.md)|
|[Parent](./Project.Subprojects.Parent.md)|

## Examples

### Using the Subprojects collection object

Use **Subprojects** (_index_), where _index_ is the subproject index or project summary task name, to return a single **Subproject** object. The following example prevents changes made to the specified subproject in a master project from being automatically made to the source project.

```vb
ActiveProject.Subprojects("Arcadia Bay Online Catalog Plan").LinkToSource = False
```

### Getting the Subprojects collection object

Use the **[Subprojects](./Project.Project.Subprojects.md)** property to return a **Subprojects** collection. The following example cautions the user if any of the subprojects in the active project are not on the hard disk.

```vb
Dim SubProj As Subproject 

For Each SubProj in ActiveProject.Subprojects 

 If UCase(Left$(SubProj.Path, 1)) <> "C" Then 

 MsgBox Right$(SubProj.Path, InStrRev(SubProj.Path, "\") - 1) & _ 

 " is not on your local hard disk.", vbExclamation 

 End If 

Next SubProj
```

> [!NOTE] 
> If you add two subprojects with the same name to a project, it will become a static object, and will not provide information about any additional subprojects that are added to your project. This continues for the life of the project file, even if one of the similarly named subprojects is removed. 
> 
> You can try this by making a copy of one of the subprojects in your project, placing it into another folder, and then adding it to your project again. Subprojects will not report the new project, or provide information about any subprojects that are added to that project afterwards.

## See also

- [Project Object Model](../project/Concepts/project-object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]