---
title: DocumentWindows.Arrange method (PowerPoint)
keywords: vbapp10.chm509004
f1_keywords:
- vbapp10.chm509004
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindows.Arrange
ms.assetid: e816fc32-e26f-6a3a-8299-7db24588778f
ms.date: 06/08/2017
localization_priority: Normal
---


# DocumentWindows.Arrange method (PowerPoint)

Arranges all open document windows in the workspace.


## Syntax

_expression_. `Arrange`( `_arrangeStyle_` )

_expression_ A variable that represents a [DocumentWindows](PowerPoint.DocumentWindows.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _arrangeStyle_|Optional|**[PpArrangeStyle](PowerPoint.PpArrangeStyle.md)**|Specifies whether to cascade or tile the windows.|

## Return value

Nothing


## Example

This example creates a new window and then arranges all open document windows.


```vb
Application.ActiveWindow.NewWindow

Windows.Arrange ppArrangeCascade
```


## See also


[DocumentWindows Object](PowerPoint.DocumentWindows.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]