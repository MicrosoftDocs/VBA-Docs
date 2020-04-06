---
title: Application.Parent property (Project)
ms.prod: project-server
api_name:
- Project.Application.Parent
ms.assetid: 4942313c-4f03-362f-0fbb-9596050a7231
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Parent property (Project)

Gets the parent of the  **Application** object. Read-only **Application**.


## Syntax

_expression_.**Parent**

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Remarks

The parent of the  **Application** object is the **Application** object.


## Example

For example, executing either of the following statements in the  **Immediate** pane of the VBE shows the text **Microsoft Project**.


```vb
? Application.Parent.Name 
? Application.Name
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]