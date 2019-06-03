---
title: Application.ProtectedViewWindowActivate event (Word)
ms.prod: word
api_name:
- Word.Application.ProtectedViewWindowActivate
ms.assetid: ae68e1aa-7cec-cd76-ee0e-71a051c5b6e3
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ProtectedViewWindowActivate event (Word)

Occurs when any Protected View window is activated.


## Syntax

_expression_. `ProtectedViewWindowActivate`( `_PvWindow_` , )

 _expression_ An expression that returns a '[Application](Word.Application.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PvWindow_|Required| **[ProtectedViewWindow](Word.ProtectedViewWindow.md)**|The Protected View window that is activated.|

## Example

The following code example maximizes any Protected View window when it is activated. This code must be placed in a class module, and an instance of the class must be correctly initialized to see this example work. For more information about how to do this, see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md).

The following code example assumes that you have declared an application variable called "App" in your general declarations and have set the variable equal to the Word Application object.




```vb
Private Sub App_ProtectedViewWindowActivate(ByVal PvWindow As ProtectedViewWindow) 
 PvWindow.WindowState = wdWindowStateMaximize 
End Sub
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]