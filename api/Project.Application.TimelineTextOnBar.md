---
title: Application.TimelineTextOnBar method (Project)
keywords: vbapj.chm63
f1_keywords:
- vbapj.chm63
ms.prod: project-server
api_name:
- Project.Application.TimelineTextOnBar
ms.assetid: d57ec0d8-8e35-b6eb-1932-454210bc7dad
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.TimelineTextOnBar method (Project)

Changes the format of text to display as a callout or within the Timeline bar, for one or more selected tasks.


## Syntax

_expression_. `TimelineTextOnBar`( `_TextOnBar_` )

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _TextOnBar_|Optional|**Boolean**|**False** if the selected tasks should be displayed as callouts; otherwise, **True**. The default value is **True**, which makes the task text show within the Timeline bar.|

## Return value

 **Boolean**


## Remarks

The **TimelineTextOnBar** method is equivalent to the **Display as Bar** and **Display as Callout** commands in the **Current Selection** group on the **Format** tab on the ribbon.


## Example

The following statement changes selected tasks on the Timeline bar to display as callouts.


```vb
TimelineTextOnBar TextOnBar:=False
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]