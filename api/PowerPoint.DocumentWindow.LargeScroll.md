---
title: DocumentWindow.LargeScroll method (PowerPoint)
keywords: vbapp10.chm511017
f1_keywords:
- vbapp10.chm511017
ms.prod: powerpoint
api_name:
- PowerPoint.DocumentWindow.LargeScroll
ms.assetid: b74ecd74-acec-0d36-68c7-1848a99fe4c1
ms.date: 06/08/2017
localization_priority: Normal
---


# DocumentWindow.LargeScroll method (PowerPoint)

Scrolls through the specified document window by pages.


## Syntax

_expression_.**LargeScroll** (_Down_, _Up_, _ToRight_, _ToLeft_)

_expression_ A variable that represents a [DocumentWindow](PowerPoint.DocumentWindow.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Down_|Optional|**Long**|Specifies the number of pages to scroll down.|
| _Up_|Optional|**Long**|Specifies the number of pages to scroll up.|
| _ToRight_|Optional|**Long**|Specifies the number of pages to scroll right.|
| _ToLeft_|Optional|**Long**|Specifies the number of pages to scroll left.|

## Remarks

If no arguments are specified, this method scrolls down one page. If Down and Up are both specified, their effects are combined. For example, if Down is 2 and Up is 4, this method scrolls up two pages. Similarly, if ToRight and ToLeft are both specified, their effects are combined.

Any of the arguments can be a negative number.


## Example

This example scrolls the active window down three pages.


```vb
Application.ActiveWindow.LargeScroll Down:=3
```


## See also


[DocumentWindow Object](PowerPoint.DocumentWindow.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]