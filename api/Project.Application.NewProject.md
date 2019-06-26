---
title: Application.NewProject event (Project)
ms.prod: project-server
api_name:
- Project.Application.NewProject
ms.assetid: de3c9e06-405a-8f63-6210-013f5d292c20
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.NewProject event (Project)

Occurs when a new project is created, including the default project that is created each time Project starts.


## Syntax

_expression_. `NewProject`( `_pj_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _pj_|Required|**Project**| The project that was created.|

## Remarks

The  **NewProject** event for the default project is analogous to the **Open** event for existing projects. The **NewProject** event occurs before the **Activate** event for a new project. Project events do not occur when the project is embedded in another document or application. For more information and sample code for creating and testing an event handler, see [Using Events with Application and Project Objects](../project/Concepts/using-events-with-application-and-project-objects.md).


## Example

The following example sets the number of working hours per day for every new project created. This example requires a new class module and additional code for it to have an effect. 


```vb
Private Sub App_NewProject(ByVal pj As MSProject.Project) 
    pj.HoursPerDay = 10 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]