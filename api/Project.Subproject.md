---
title: Subproject object (Project)
ms.prod: project-server
api_name:
- Project.Subproject
ms.assetid: 1a3b0d18-6464-a4f2-479f-710e19faffa8
ms.date: 06/08/2017
localization_priority: Normal
---


# Subproject object (Project)



Represents a subproject. The  **Subproject** object is a member of the **[Subprojects](Project.subprojects(object).md)** collection.
 **Using the Subproject Object**
Use  **Subprojects** (_index_), where _index_ is the subproject index or project summary task name, to return a single **Subproject** object. The following example prevents changes made to the specified subproject in a master project from being automatically made to the source project.
 **Using the Subprojects Collection**
Use the  **[Subprojects](./Project.Project.Subprojects.md)** property to return a **Subprojects** collection. The following example cautions the user if any of the subprojects in the active project are not on the hard disk.

## Properties



|Name|
|:-----|
|[Application](./Project.Subproject.Application.md)|
|[Index](./Project.Subproject.Index.md)|
|[InsertedProjectSummary](./Project.Subproject.InsertedProjectSummary.md)|
|[IsLoaded](./Project.Subproject.IsLoaded.md)|
|[LinkToSource](./Project.Subproject.LinkToSource.md)|
|[Parent](./Project.Subproject.Parent.md)|
|[Path](./Project.Subproject.Path.md)|
|[ReadOnly](./Project.Subproject.ReadOnly.md)|
|[SourceProject](./Project.Subproject.SourceProject.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]