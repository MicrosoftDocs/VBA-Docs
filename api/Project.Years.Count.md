---
title: Years.Count property (Project)
ms.prod: project-server
api_name:
- Project.Years.Count
ms.assetid: 6a65ff7b-55ca-31e0-0edd-c2f75cb9fc74
ms.date: 06/08/2017
localization_priority: Normal
---


# Years.Count property (Project)

Gets the number of items in the **Years** collection. Read-only **Integer**.


## Syntax

_expression_.**Count**

_expression_ A variable that represents a 'Years' object.


## Remarks

The following statement prints 166 in the **Immediate** pane of the VBE. The value is the number of years from 1984 to and including 2149.


```vb
Print ActiveProject.Calendar.Years.Count
```

Use of the **Count** property in most collection objects is similar. For an example that uses the **Years** collection, see [Years Object](Project.years.md).


## See also


[Years Collection Object](Project.years.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]