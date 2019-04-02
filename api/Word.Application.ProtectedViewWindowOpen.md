---
title: Application.ProtectedViewWindowOpen event (Word)
keywords: vbawd10.chm4000030
f1_keywords:
- vbawd10.chm4000030
ms.prod: word
api_name:
- Word.Application.ProtectedViewWindowOpen
ms.assetid: 42126a64-0227-d006-760e-ec11c59ef533
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ProtectedViewWindowOpen event (Word)

Occurs when a protected view window is opened.


## Syntax

_expression_. `ProtectedViewWindowOpen`( `_PvWindow_` , )

 _expression_ An expression that returns a '[Application](Word.Application.md)' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PvWindow_|Required| **[ProtectedViewWindow](Word.ProtectedViewWindow.md)**|The protected view window that is opened.|

## Example

The following code example informs the user that the document will be opened in a protected view window. This code must be placed in a class module, and an instance of the class must be correctly initialized for this code example to work correctly. For more information about how to do this, see [Using Events with the Application Object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md).

The following code example assumes that you have declared an application variable called "App" in your general declarations and have set the variable equal to the Word Application object.




```vb
Private Sub App_ProtectedViewWindowOpen(ByVal PvWindow As ProtectedViewWindow) 
Dim intResponse As Integer 
 
 MsgBox "You are opening a document in " _ 
 & "protected view window mode." 
 
End Sub
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]