---
title: Application.ActiveProject property (Project)
keywords: vbapj.chm131377
f1_keywords:
- vbapj.chm131377
ms.prod: project-server
api_name:
- Project.Application.ActiveProject
ms.assetid: 07844166-ca9b-15eb-a5e2-6f00a7c0a030
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ActiveProject property (Project)

Gets a  **[Project](Project.Project.md)** object that represents the active project. Read-only **Project**.


## Syntax

_expression_. `ActiveProject`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Example

The following example adds the date and time to the  **Comments** field in the project **Properties** dialog box and then saves the project.


```vb
Sub SaveAndNoteTime() 
 ActiveProject.ProjectNotes = ActiveProject.ProjectNotes & vbCrLf _ 
 & "This project was last saved on " & Date$ & " at " & Time$ & "." 
 FileSave 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]