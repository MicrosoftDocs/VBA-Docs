---
title: Application.DisplayScrollBars property (Project)
ms.prod: project-server
api_name:
- Project.Application.DisplayScrollBars
ms.assetid: 4c8e2aa3-3d85-94c8-d1ce-67586b78e7e7
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.DisplayScrollBars property (Project)

 **True** if the scroll bars are visible for all projects. Read/write **Boolean**.


## Syntax

_expression_. `DisplayScrollBars`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Example

The following example changes the setting of the  **DisplayScrollBars** property.


```vb
Sub ChangeDisplayScrollBars 
 DisplayScrollBars = Not DisplayScrollBars 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]