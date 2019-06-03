---
title: Application.WindowDeactivate event (Word)
keywords: vbawd10.chm4000010
f1_keywords:
- vbawd10.chm4000010
ms.prod: word
api_name:
- Word.Application.WindowDeactivate
ms.assetid: 70b86ecc-40ba-6e70-b430-4fce6083ff2d
ms.date: 08/20/2018
localization_priority: Normal
---


# Application.WindowDeactivate event (Word)

Occurs when any document window is deactivated.

> [!NOTE] 
> If you are working with a document embedded within another document, this event will not occur.

## Syntax

_expression_. `Private Sub object_WindowDeactivate`(`ByVal Doc As Document`, `ByVal Wn As Window`)

_expression_ A variable that represents an [Application](Word.Application.md) object that has been declared with events in a class module. For more information about using events with the **Application** object or the **Document** object, see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md) or [Using Events with the Document Object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-document-object.md).


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The document displayed in the deactivated window.|
| _Wn_|Required| **Window**|The deactivated window.|

## Example

This example minimizes any document window when it is deactivated. This code must be placed in a class module, and an instance of the class must be correctly initialized to see this example work; see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md) for directions on how to accomplish this.


```vb
Public WithEvents appWord as Word.Application 
 
Private Sub appWord_WindowDeactivate _ 
 (ByVal Wn As Word.Window) 
 Wn.WindowState = wdWindowStateMinimize 
End Sub
```


## See also

- [Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]