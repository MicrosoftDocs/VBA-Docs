---
title: Project.CreationDate property (Project)
keywords: vbapj.chm131693
f1_keywords:
- vbapj.chm131693
ms.prod: project-server
api_name:
- Project.Project.CreationDate
ms.assetid: 7126f72b-fe35-c183-04b7-03efd78a8589
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.CreationDate property (Project)

Gets the date a project was created. Read-only  **Variant**.


## Syntax

_expression_. `CreationDate`

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Example

The following example adds the creation date of the active project to its notes.


```vb
Sub AddCreationDateToNotes() 
 ActiveProject.ProjectNotes = ActiveProject.ProjectNotes & vbCrLf & "This project was created on " & ActiveProject.CreationDate & "." 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]