---
title: Application.WindowActivate event (Word)
keywords: vbawd10.chm400009
f1_keywords:
- vbawd10.chm400009
ms.prod: word
api_name:
- Word.Application.WindowActivate
ms.assetid: f1340e1e-6aec-edaa-78c2-47e3e1d5299f
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.WindowActivate event (Word)

Occurs when any document window is activated.


## Syntax

_expression_.**WindowActivate** (_Doc_, _Wn_)

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 

For more information about using events with the **Application** object, see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md).


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The document displayed in the activated window.|
| _Wn_|Required| **Window**|The window that's being activated.|

## Example

This example maximizes any document window when it is activated. This code must be placed in a class module, and an instance of the class must be correctly initialized to see this example work; see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md)for directions on how to accomplish this.


```vb
Public WithEvents appWord as Word.Application 
 
Private Sub appWord_WindowActivate _ 
 (ByVal Doc As Word.Document, _
  ByVal Wn As Word.Window) 
 Wn.WindowState = wdWindowStateMaximize 
End Sub
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]