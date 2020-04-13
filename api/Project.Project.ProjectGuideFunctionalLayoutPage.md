---
title: Project.ProjectGuideFunctionalLayoutPage property (Project)
keywords: vbapj.chm131088
f1_keywords:
- vbapj.chm131088
ms.prod: project-server
api_name:
- Project.Project.ProjectGuideFunctionalLayoutPage
ms.assetid: 87a9e383-6b91-669e-86e4-e55b7030b861
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.ProjectGuideFunctionalLayoutPage property (Project)

Gets or sets the Project Guide functional layout page for the specified project. Read/write  **String**.


## Syntax

_expression_. `ProjectGuideFunctionalLayoutPage`

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Remarks


> [!NOTE] 
> The Project Guide is deprecated in Project. Instead of the Project Guide, we recommend that you create task pane apps.

However, you can still use custom Project Guides and get the default Project Guide files from the Project SDK download. The Project Guide files are modified for access in a flat folder structure and to remove the  `gbui://` protocol (**gbui** is the goal-based user interface protocol in Office Project 2007 and previous versions). All Project Guide settings must be made programmatically.

The default value of the **ProjectGuideFunctionalLayoutPage** property is `gbui://mainpage.htm`, which does not work because Project does not implement the  `gbui://` protocol. The Project Programmability blog ( `https://blogs.msdn.com/project_programmability/`) includes articles that show how to use the Project Guide in a VBA macro and in an add-in that is developed with Visual C# in Microsoft Office development tools in Visual Studio 2010.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]