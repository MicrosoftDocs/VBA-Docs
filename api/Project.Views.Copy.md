---
title: Views.Copy method (Project)
ms.prod: project-server
api_name:
- Project.Views.Copy
ms.assetid: 5e82641a-f5c6-41a6-23bf-61220a4fc30c
ms.date: 06/08/2017
localization_priority: Normal
---


# Views.Copy method (Project)

Makes a copy of a group definition for the  **Views** collection and returns a reference to the **[View](Project.View.md)** object.


## Syntax

_expression_.**Copy** (_Source_, _NewName_)

_expression_ A variable that represents a 'Views' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Source_|Required|**String**|The name of the view to copy.|
| _NewName_|Required|**String**|The name of the new view.|

## Return value

 **View**


## See also


[Views Collection Object](Project.views(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]