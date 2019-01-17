---
title: Pane.LargeScroll method (Word)
keywords: vbawd10.chm157286502
f1_keywords:
- vbawd10.chm157286502
ms.prod: word
api_name:
- Word.Pane.LargeScroll
ms.assetid: 669093ef-105f-5dc2-19d6-5f8ab11f6d50
ms.date: 06/08/2017
localization_priority: Normal
---


# Pane.LargeScroll method (Word)

Scrolls a window or pane by the specified number of screens.


## Syntax

 _expression_. `LargeScroll`( `_Down_` , `_Up_` , `_ToRight_` , `_ToLeft_` )

 _expression_ Required. A variable that represents a '[Pane](Word.Pane.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Down_|Optional| **Variant**|The number of screens to scroll the window down.|
| _Up_|Optional| **Variant**|The number of screens to scroll the window up.|
| _ToRight_|Optional| **Variant**|The number of screens to scroll the window to the right.|
| _ToLeft_|Optional| **Variant**|The number of screens to scroll the window to the left.|

## Remarks

This method is equivalent to clicking just before or just after the scroll boxes on the horizontal and vertical scroll bars.

If Down and Up are both specified, the window is scrolled by the difference of the arguments. For example, if Down is 2 and Up is 4, the window is scrolled up two screens. Similarly, if ToLeft and ToRight are both specified, the window is scrolled by the difference of the arguments.

Any of these arguments can be a negative number. If no arguments are specified, the window is scrolled down one screen.


## See also


[Pane Object](Word.Pane.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]