---
title: Panes.Add method (Word)
keywords: vbawd10.chm157220867
f1_keywords:
- vbawd10.chm157220867
ms.prod: word
api_name:
- Word.Panes.Add
ms.assetid: 34dba7e0-cb4f-0482-c8c5-cc3d54cacc9c
ms.date: 06/08/2017
localization_priority: Normal
---


# Panes.Add method (Word)

Returns a  **Pane** object that represents a new pane to a window.


## Syntax

 _expression_. `Add`( `_SplitVertical_` )

 _expression_ Required. A variable that represents a '[Panes](Word.panes.md)' collection.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SplitVertical_|Optional| **Variant**|A number that represents the percentage of the window, from top to bottom, you want to appear above the split.|

## Return value

Pane


## Remarks

This method will fail if it is applied to a window that has already been split.


## Example

The following example splits the active window such that the top pane is 30 percent of the total window size.


```vb
ActiveDocument.ActiveWindow.Panes.Add SplitVertical:=30
```


## See also


[Panes Collection Object](Word.panes.md)

