---
title: DocumentWindow.WindowState Property (PowerPoint)
keywords: vbapp10.chm511009
f1_keywords:
- vbapp10.chm511009
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindow.WindowState
ms.assetid: 7f0ce168-0339-03f0-11e4-dc7935c04b85
ms.date: 06/08/2017
localization_priority: Normal
---


# DocumentWindow.WindowState Property (PowerPoint)

Returns or sets the state of the specified window. Read/write.


## Syntax

 _expression_. `WindowState`

_expression_ A variable that represents a [DocumentWindow](./PowerPoint.DocumentWindow.md) object.


## Return value

PpWindowState


## Remarks

The value of the  **WindowState** property can be one of these **PpWindowState** constants.


||
|:-----|
|**ppWindowMaximized**|
|**ppWindowMinimized**|
|**ppWindowNormal**|

When the state of the window is  **ppWindowNormal**, the window is neither maximized nor minimized.


## Example

This example maximizes the first member of the  **DocumentWindows** collection.


```vb
Windows(1).WindowState = ppWindowMaximized
```


## See also



[DocumentWindow Object](PowerPoint.DocumentWindow.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]