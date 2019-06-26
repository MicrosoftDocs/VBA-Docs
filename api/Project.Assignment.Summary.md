---
title: Assignment.Summary property (Project)
ms.prod: project-server
api_name:
- Project.Assignment.Summary
ms.assetid: 7f8f38f3-c712-0f4e-6b46-0d8eb02119f4
ms.date: 06/08/2017
localization_priority: Normal
---


# Assignment.Summary property (Project)

Indicates whether the assignment is on a summary task. Read-only **String**.


## Syntax

_expression_.**Summary**

_expression_ A variable that represents an [Assignment](./Project.Assignment.md) object.


## Remarks

For an example that checks whether summary tasks have assignments, see the **[Summary](Project.Task.Summary.md)** property for the **Task** object.


> [!NOTE] 
> Project ignores the **Summary** property for an assignment. The property value is always "No".

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]