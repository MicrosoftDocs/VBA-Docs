---
title: Panes.Item property (Excel)
keywords: vbaxl10.chm358074
f1_keywords:
- vbaxl10.chm358074
ms.prod: excel
api_name:
- Excel.Panes.Item
ms.assetid: 5960e77c-23b4-2ce4-1674-2ffd3b4f6e47
ms.date: 05/03/2019
localization_priority: Normal
---


# Panes.Item property (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[Panes](Excel.Panes.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index number of the object.|

## Example

This example splits the window in which worksheet one is displayed, and then scrolls through the pane in the lower-left corner of the window until row five is at the top of the pane.

```vb
Worksheets(1).Activate 
ActiveWindow.Split = True 
ActiveWindow.Panes.Item(3).ScrollRow = 5
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]