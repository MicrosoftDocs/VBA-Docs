---
title: Filters object (Project)
ms.prod: project-server
ms.assetid: 13b58540-decc-17c5-6de6-bbb8e05eb6d2
ms.date: 06/08/2017
localization_priority: Normal
---


# Filters object (Project)

Contains a collection of  **[Filter](Project.Filter.md)** objects.
 


## Example

 **Using the Filters Collection**
 

 
The following example applies a critical task filter to the active project. 
 

 



```vb
ActiveProject.TaskFilters("Critical").Apply
```


## Methods



|Name|
|:-----|
|[Copy](Project.Filters.Copy.md)|

## Properties



|Name|
|:-----|
|[Application](Project.Filters.Application.md)|
|[Count](Project.Filters.Count.md)|
|[Item](Project.Filters.Item.md)|
|[Parent](Project.Filters.Parent.md)|

## See also


 
[Project Object Model](../project/Concepts/project-object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]