---
title: Application.Undo method (Project)
keywords: vbapj.chm132718
f1_keywords:
- vbapj.chm132718
ms.prod: project-server
api_name:
- Project.Application.Undo
ms.assetid: 50e1b5ba-fe4b-d53d-5712-8e2023eb2755
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Undo method (Project)

Executes an undo action on items in the  **Undo** list.


## Syntax

_expression_.**Undo**( `_HowManyUndos_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _HowManyUndos_|Optional|**Long**|Specifies the number of items from the list to undo. The default is 1.|

## Return value

 **Boolean**


## Remarks

Many actions you perform in Project, such as adding a task, add items to the  **Undo** list. To redo one or more actions after using the **Undo** method, you can use the **[Redo](Project.Application.Redo.md)** method or click **Redo** in the Quick Access Toolbar.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]