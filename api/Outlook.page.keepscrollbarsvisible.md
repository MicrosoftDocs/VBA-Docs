---
title: Page.KeepScrollBarsVisible Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 4abf7176-4460-91b6-03e1-291b71db0752
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.KeepScrollBarsVisible Property (Outlook Forms Script)

Returns or sets an **Integer** that specifies whether scroll bars remain visible when not required. Read/write.


## Syntax

_expression_.**KeepScrollBarsVisible**

_expression_ A variable that represents a **Page** object.


## Remarks

The settings for  **KeepScrollBarsVisible** are:



|Value|Description|
|:-----|:-----|
|0|Displays no scroll bars.|
|1|Displays a horizontal scroll bar.|
|2|Displays a vertical scroll bar.|
|3|Displays both a horizontal and a vertical scroll bar (default).|

If the visible region is large enough to display all the controls on an object such as a **[Page](Outlook.page.md)** object, scroll bars are not required. The **KeepScrollBarsVisible** property determines whether the scroll bars remain visible when they are not required.

If the scroll bars are visible when they are not required, they appear normal in size, and the scroll box fills the entire width or height of the scroll bar.

If the  **KeepScrollBarsVisible** property is **True**, any scroll bar on a form or page is always visible, regardless of whether the object's contents fit within the object's borders.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]