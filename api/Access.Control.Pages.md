---
title: Control.Pages property (Access)
keywords: vbaac10.chm10149
f1_keywords:
- vbaac10.chm10149
ms.prod: access
api_name:
- Access.Control.Pages
ms.assetid: fd4ea2c0-ea8c-51a0-a012-8ba5848d3516
ms.date: 06/08/2017
localization_priority: Normal
---


# Control.Pages property (Access)

Returns a  **[Pages](Access.Pages.md)** collection that represents the pages in the specified control that supports tabbed pages (for example, a **TabControl** object). Read-only.


## Syntax

_expression_. `Pages`

_expression_ A variable that represents a [Control](Access.Control.md) object.


## Example

The following example displays a message indicating the number of tabbed pages on tab control TabCtl1.


```vb
MsgBox "Number of pages in TabCtl1:" & TabCtl1.Pages.Count
```


## See also


[Control Object](Access.Control.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]