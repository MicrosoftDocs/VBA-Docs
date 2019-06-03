---
title: Application.ProtectedViewWindowDeactivate event (Word)
keywords: vbawd10.chm4000035
f1_keywords:
- vbawd10.chm4000035
ms.prod: word
api_name:
- Word.Application.ProtectedViewWindowDeactivate
ms.assetid: bd80056b-edce-7e0b-c61a-31ebda24a416
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ProtectedViewWindowDeactivate event (Word)

Occurs when a Protected View window is deactivated.


## Syntax

_expression_. `ProtectedViewWindowDeactivate`( `_PvWindow_` , )

 _expression_ An expression that returns a '[Application](Word.Application.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PvWindow_|Required| **[ProtectedViewWindow](Word.ProtectedViewWindow.md)**|The deactivated Protected View window.|

## Example

The following code example minimizes an open Protected View window when it is deactivated. This code must be placed in a class module, and an instance of the class must be correctly initialized for this code example to work correctly. For more information about how to do this, see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md).

The following code example assumes that you have declared an application variable called "App" in your general declarations and have set the variable equal to the Word Application object.




```vb
Private Sub App_ProtectedViewWindowDeactivate(ByVal PvWindow As ProtectedViewWindow) 
 PvWindow.WindowState = wdWindowStateMinimize 
End Sub
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]