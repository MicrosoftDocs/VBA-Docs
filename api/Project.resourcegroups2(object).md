---
title: ResourceGroups2 object (Project)
ms.prod: project-server
ms.assetid: b1328c39-42bc-4e9b-e268-1f308cd7ebb1
ms.date: 06/08/2017
localization_priority: Normal
---


# ResourceGroups2 object (Project)

Represents all of the resource-based group definitions, where group hierarchy can be maintained.  **ResourceGroups2** is a collection of **[Group2](Project.Group2.md)** objects.
 


## Example

 **Using the ResourceGroups2 Collection**
 

 
Use the **[ResourceGroups2](Project.Project.ResourceGroups2.md)** property to return a **ResourceGroups2** collection. The following example lists the names of all the resource groups in the active project.
 

 



```vb
Dim rg2 As Group2  
Dim rGroups2 As String  
  
For Each rg2 in ActiveProject.ResourceGroups2  
    rGroups2 = rGroups2 & rg2.Name & vbCrLf  
Next rg2  
  
MsgBox rGroups2
```

Use the **[Add](Project.ResourceGroups2.Add.md)** method to add a **Group2** object to the **ResourceGroups2** collection. The following example creates a new group that groups resources by their standard rate and then modifies the criterion so that the resources are sorted in descending order.
 

 



```vb
ActiveProject.ResourceGroups2.Add "Resources by Rate", "Standard Rate"  
ActiveProject.ResourceGroups2("Resources by Rate").GroupCriteria(1).Ascending = False
```


## Methods



|Name|
|:-----|
|[Add](Project.ResourceGroups2.Add.md)|
|[Copy](Project.ResourceGroups2.Copy.md)|

## Properties



|Name|
|:-----|
|[Application](Project.ResourceGroups2.Application.md)|
|[Count](Project.ResourceGroups2.Count.md)|
|[Item](Project.ResourceGroups2.Item.md)|
|[Parent](Project.ResourceGroups2.Parent.md)|

## See also


 
[Project Object Model](../project/Concepts/project-object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]