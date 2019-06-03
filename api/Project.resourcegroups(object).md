---
title: ResourceGroups object (Project)
ms.prod: project-server
ms.assetid: 37bd0f3a-4d0e-1311-4409-ed31e0fe2e3a
ms.date: 06/08/2017
localization_priority: Normal
---


# ResourceGroups object (Project)


 

Represents all of the resource-based group definitions.  **ResourceGroups** is a collection of **[Group](Project.Group.md)** objects.
 
 **Using the ResourceGroups Collection**
 
Use the  **[ResourceGroups](Project.Project.ResourceGroups.md)** property to return a **ResourceGroups** collection. The following example lists the names of all the resource groups in the active project.
 



```vb
Dim rg As Group 
Dim rGroups As String 
 
For Each rg in ActiveProject.ResourceGroups 
 rGroups = rGroups & rg.Name & vbCrLf 
Next rg 
 
MsgBox rGroups
```

Use the  **[Add](Project.ResourceGroups.Add.md)** method to add a **Group** object to the **ResourceGroups** collection. The following example creates a new group that groups resources by their standard rate and then modifies the criterion so that the resources are sorted in descending order.
 



```vb
ActiveProject.ResourceGroups.Add "Resources by Rate", "Standard Rate" 
ActiveProject.ResourceGroups("Resources by Rate").GroupCriteria(1).Ascending = False
```


## Remarks

For resource groups where the group hierarchy can be maintained and cell color can be a hexadecimal value, use the  **[ResourceGroups2](Project.resourcegroups2(object).md)** collection object.
 

 

## Methods



|Name|
|:-----|
|[Add](Project.ResourceGroups.Add.md)|
|[Copy](Project.ResourceGroups.Copy.md)|

## Properties



|Name|
|:-----|
|[Application](Project.ResourceGroups.Application.md)|
|[Count](Project.ResourceGroups.Count.md)|
|[Item](Project.ResourceGroups.Item.md)|
|[Parent](Project.ResourceGroups.Parent.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]