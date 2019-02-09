---
title: Document.Undo method (Word)
keywords: vbawd10.chm158007412
f1_keywords:
- vbawd10.chm158007412
ms.prod: word
api_name:
- Word.Document.Undo
ms.assetid: f9fd64c9-aeb9-b698-6318-beb1db653ee6
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.Undo method (Word)

Undoes the last action or a sequence of actions, which are displayed in the **Undo** list. Returns **True** if the actions were successfully undone.


## Syntax

_expression_.**Undo** (_Times_)

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Times_|Optional| **Variant**|The number of actions to be undone.|

## Return value

Boolean


## Example

This example undoes the last two actions taken in Sales.doc.


```vb
Documents("Sales.doc").Undo 2
```

This example undoes the last action. If the action is successfully undone, a message is displayed in the status bar.




```vb
On Error Resume Next 
If ActiveDocument.Undo = False Then _ 
 StatusBar = "Undo was unsuccessful"
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]