---
title: Pane.View property (Word)
keywords: vbawd10.chm157286410
f1_keywords:
- vbawd10.chm157286410
ms.prod: word
api_name:
- Word.Pane.View
ms.assetid: 64e4d06a-8e4e-ce65-1732-66865eff4125
ms.date: 06/08/2017
localization_priority: Normal
---


# Pane.View property (Word)

Returns a  **View** object that represents the view for the specified pane.


## Syntax

_expression_.**View**

_expression_ Required. A variable that represents a '[Pane](Word.Pane.md)' object.


## Example

This example shows all nonprinting characters for panes associated with the first window in the **Windows** collection.


```vb
For Each myPane In Windows(1).Panes 
 myPane.View.ShowAll = True 
Next myPane
```


## See also


[Pane Object](Word.Pane.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]