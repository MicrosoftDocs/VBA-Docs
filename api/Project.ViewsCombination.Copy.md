---
title: ViewsCombination.Copy method (Project)
ms.prod: project-server
api_name:
- Project.ViewsCombination.Copy
ms.assetid: 2e28885e-6b65-8123-193a-1ac0ee883f75
ms.date: 06/08/2017
localization_priority: Normal
---


# ViewsCombination.Copy method (Project)

Makes a copy of a group definition for the  **ViewsCombination** collection and returns a reference to the **[View](Project.View.md)** object.


## Syntax

_expression_.**Copy** (_Source_, _NewName_)

_expression_ A variable that represents a 'ViewsCombination' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Source_|Required|**String**|The name of the view to copy.|
| _NewName_|Required|**String**|The name of the new view.|

## Return value

 **View**


## See also


[ViewsCombination Collection Object](Project.viewscombination(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]