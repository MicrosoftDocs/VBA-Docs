---
title: Calendar.Name property (Project)
ms.prod: project-server
api_name:
- Project.Calendar.Name
ms.assetid: e437e29c-ed61-c83a-53b7-8a0d1cb7cb4e
ms.date: 06/08/2017
localization_priority: Normal
---


# Calendar.Name property (Project)

Gets the name of a **Calendar** object. Read-only **String**.


## Syntax

_expression_.**Name**

_expression_ A variable that represents a [Calendar](./Project.Calendar.md) object.


## Remarks

For a code example that uses the **Task** object, see **[Name](Project.Task.Name.md)**.


## Example

 **Name** is the default property for the **Calendar** object. The following example prints the name of the default calendar for the active project.


```vb
Debug.Print ActiveProject.Calendar
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]