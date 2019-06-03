---
title: Application.ProtectedViewWindowSize event (Word)
ms.prod: word
api_name:
- Word.Application.ProtectedViewWindowSize
ms.assetid: b28d53f9-783f-6d68-2080-a0b1d8484c43
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ProtectedViewWindowSize event (Word)




## Syntax

_expression_. `ProtectedViewWindowSize`( `_PvWindow_` , )

 _expression_ An expression that returns a '[Application](Word.Application.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PvWindow_|Required| **[ProtectedViewWindow](Word.ProtectedViewWindow.md)**|The Protected View window that is sized.|

## Example

The following code example displays a message every time a Protected View window is moved or resized. This code must be placed in a class module, and an instance of the class must be correctly initialized for this code example to work correctly. For more information about how to do this, see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md).

The following code example assumes that you have declared an application variable called "App" in your general declarations and have set the variable equal to the Word Application object.




```vb
Private Sub App_ProtectedViewWindowSize(ByVal PvWindow As ProtectedViewWindow) 
MsgBox "You resized a window!" 
End Sub
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]