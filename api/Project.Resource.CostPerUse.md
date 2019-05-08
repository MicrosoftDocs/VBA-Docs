---
title: Resource.CostPerUse property (Project)
ms.prod: project-server
api_name:
- Project.Resource.CostPerUse
ms.assetid: 171217c9-200b-8cd1-b985-aa1aed099d0e
ms.date: 06/08/2017
localization_priority: Normal
---


# Resource.CostPerUse property (Project)

Gets or sets the cost per use of a resource. Read/write  **Variant**.


## Syntax

_expression_. `CostPerUse`

_expression_ A variable that represents a [Resource](./Project.Resource.md) object.


## Example

The following example displays the sum of the cost per use of each resource in the active project.


```vb
Sub TotalCostPerUse() 
 
 Dim R As Resource ' Resource object used in For Each loop 
 Dim TotalCostPerUse As Double ' The total cost per use 
 
 ' Add up the cost per use of each resource. 
 For Each R In ActiveProject.Resources 
 TotalCostPerUse = TotalCostPerUse + R.CostPerUse 
 Next R 
 
 ' Display the total cost per use. 
 MsgBox ("Sum of the cost per use of each resource in this project: " & TotalCostPerUse) 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]