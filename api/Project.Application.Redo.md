---
title: Application.Redo method (Project)
keywords: vbapj.chm132540
f1_keywords:
- vbapj.chm132540
ms.prod: project-server
api_name:
- Project.Application.Redo
ms.assetid: 25a43bd7-4bfd-2be6-172d-8e5bef781f00
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Redo method (Project)

Executes a redo action on items in the **Redo** list.


## Syntax

_expression_.**Redo** (_HowManyRedos_)

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _HowManyRedos_|Optional|**Long**|Specifies the number of items from the list to redo. The default is 1.|

## Return value

 **Boolean**


## Remarks

You can add items to the **Redo** list by using the **[Undo](Project.Application.Undo.md)** method or clicking **Undo** in the Quick Access Toolbar.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]