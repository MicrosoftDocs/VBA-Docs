---
title: CustomTaskPane.DockPosition property (Office)
keywords: vbaof11.chm301008
f1_keywords:
- vbaof11.chm301008
ms.prod: office
api_name:
- Office.CustomTaskPane.DockPosition
ms.assetid: 591c3f81-545f-6b04-7c4c-a3a85946e161
ms.date: 01/04/2019
localization_priority: Normal
---


# CustomTaskPane.DockPosition property (Office)

Gets or sets an enumerated value specifying the docked position of a **CustomTaskPane** object. Read/write.


## Syntax

_expression_.**DockPosition**

_expression_ An expression that returns a **[CustomTaskPane](Office.CustomTaskPane.md)** object.


## Return value

MsoCTPDockPosition


## Remarks

Defaults to **Right** for right-to-left languages and **Left** for left-to-right languages.

The value of this property can be set to one of the following **[MsoCTPDockPosition](office.msoctpdockposition.md)** constants.

|Name|Value|Description|
|:-----|:-----|:-----|
|**msoCTPDockPositionBottom**|3|Dock the task pane at the bottom of the document window.|
|**msoCTPDockPositionFloating**|4|Don't dock the task pane.|
|**msoCTPDockPositionLeft**|0|Dock the task pane on the left side of the document window.|
|**msoCTPDockPositionRight**|2|Dock the task pane on the right side of the document window.|
|**msoCTPDockPositionTop**|1|Dock the task pane at the top of the document window.|

## See also

- [CustomTaskPane object members](overview/library-reference/customtaskpane-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]