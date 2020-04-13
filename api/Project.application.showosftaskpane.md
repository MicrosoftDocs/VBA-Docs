---
title: Application.ShowOSFTaskPane method (Project)
keywords: vbapj.chm2199
f1_keywords:
- vbapj.chm2199
ms.prod: project-server
ms.assetid: 50109216-a0e4-ed18-ea92-e0689f896b86
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ShowOSFTaskPane method (Project)
Shows an empty Office Add-ins task pane.

## Syntax

_expression_. `ShowOSFTaskPane` _(Name)_

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the task pane app.
> [!NOTE] 
> Not implemented in Project.

|

## Return value

 **Boolean**

 **True** if the task pane is displayed; otherwise, **False**.


## Remarks

The **ShowOSFTaskPane** method is not fully implemented in Project. If another task pane app has been loaded, the **ShowOSFTaskPane** method displays an empty Office Add-ins task pane with an **APP ERROR** message. If another task pane app has not previously been loaded, the **ShowOSFTaskPane** method does nothing.


## See also


[Application Object](Project.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]