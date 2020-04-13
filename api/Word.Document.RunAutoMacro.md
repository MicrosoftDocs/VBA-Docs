---
title: Document.RunAutoMacro method (Word)
keywords: vbawd10.chm158007408
f1_keywords:
- vbawd10.chm158007408
ms.prod: word
api_name:
- Word.Document.RunAutoMacro
ms.assetid: 8eee80a6-e347-2fbb-ec86-65d09e09c764
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.RunAutoMacro method (Word)

Runs an auto macro that's stored in the specified document. If the specified auto macro doesn't exist, nothing happens.


## Syntax

_expression_. `RunAutoMacro`( `_Which_` )

_expression_ Required. A variable that represents a **[Document](Word.Document.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Which_|Required| **WdAutoMacros**|The auto macro to run.|

## Remarks

Use the **Run** method to run any macro.


## Example

This example runs the AutoOpen macro in the active document.


```vb
ActiveDocument.RunAutoMacro Which:=wdAutoOpen
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]