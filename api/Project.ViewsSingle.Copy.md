---
title: ViewsSingle.Copy method (Project)
ms.prod: project-server
api_name:
- Project.ViewsSingle.Copy
ms.assetid: baa16562-5622-6d0f-02a7-3145a6fdef0c
ms.date: 06/08/2017
localization_priority: Normal
---


# ViewsSingle.Copy method (Project)

Makes a copy of a group definition for the  **ViewsSingle** collection and returns a reference to the **[View](Project.ViewSingle.md)** object.


## Syntax

_expression_.**Copy** (_Source_, _NewName_)

_expression_ A variable that represents a 'ViewsSingle' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Source_|Required|**Variant**|The name of view to copy.|
| _NewName_|Required|**String**|The name of the new view.|

## Return value

 **View**


## See also


[ViewsSingle Collection Object](Project.viewssingle(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]