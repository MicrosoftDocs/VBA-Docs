---
title: Application.BoxZoom method (Project)
keywords: vbapj.chm308
f1_keywords:
- vbapj.chm308
ms.prod: project-server
api_name:
- Project.Application.BoxZoom
ms.assetid: fbfae092-93b1-b72f-6b42-a498a1543e00
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.BoxZoom method (Project)

Zooms in to or out from the Network Diagram.


## Syntax

_expression_. `BoxZoom`( `_Percent_`, `_Entire_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Percent_|Optional|**Variant**|The percentage&mdash;between 25 and 400&mdash;to reduce or enlarge the Network Diagram. If  **Entire** is **True**, **Percent** is ignored.|
| _Entire_|Optional|**Boolean**|**True** if the Network Diagram resizes to fit the entire project onto the screen, within the same limits described for **Percent**. The default value is **False**.|

## Return value

 **Boolean**


## Example

The following example attempts to fit all tasks onto the screen. It assumes the Network Diagram is the active view.


```vb
Sub Display() 
 BoxZoom Entire:=True 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]