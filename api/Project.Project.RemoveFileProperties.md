---
title: Project.RemoveFileProperties property (Project)
ms.prod: project-server
api_name:
- Project.Project.RemoveFileProperties
ms.assetid: 7aff624c-e9c9-f526-b233-fe0cc415e901
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.RemoveFileProperties property (Project)

 **True** if Project removes user information from revisions and the project **Properties** dialog box upon saving a document. Read/write **Boolean**.


## Syntax

_expression_. `RemoveFileProperties`

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Example

The following example sets the current document to remove personal information from File Properties in the document the next time the user saves it.


```vb
Sub RemoveFileProperties() 
 ActiveProject.RemoveFileProperties = True 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]