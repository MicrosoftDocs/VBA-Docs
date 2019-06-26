---
title: Application.OpenFromSharePoint method (Project)
keywords: vbapj.chm2293
f1_keywords:
- vbapj.chm2293
ms.prod: project-server
api_name:
- Project.Application.OpenFromSharePoint
ms.assetid: 415f8b11-5c6f-d9df-fb58-61ff7f392b5f
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.OpenFromSharePoint method (Project)

Opens a project from a task list in SharePoint 2013. 


## Syntax

_expression_. `OpenFromSharePoint`( `_SiteURL_`, `_ListName_` )

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SiteURL_|Optional|**String**|Specifies the URL of the SharePoint site.|
| _ListName_|Optional|**String**|Specifies the name of the task list.|

## Return value

 **Boolean**


## Remarks


> [!NOTE] 
> Project must not be connected to a Project Server instance. Synchronization with SharePoint task lists is designed for users who do not have access to Project Server.


## Example

The following example opens a project from a task list named TestTasks that is in the Simple project workspace.


```vb
Sub OpenSharePointTaskList() 
    OpenFromSharePoint siteurl:="https://ServerName/PWA/Simple", ListName:="TestTasks" 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]