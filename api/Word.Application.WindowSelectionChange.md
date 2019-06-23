---
title: Application.WindowSelectionChange event (Word)
keywords: vbawd10.chm4000011
f1_keywords:
- vbawd10.chm4000011
ms.prod: word
api_name:
- Word.Application.WindowSelectionChange
ms.assetid: 2c5cc640-a3a4-46b2-3352-20a057854b3a
ms.date: 08/20/2018
localization_priority: Normal
---


# Application.WindowSelectionChange event (Word)

Occurs when the selection changes in the active document window.

> [!NOTE] 
> If you are working with a document embedded within another document, this event will not occur.

## Syntax

_expression_.**WindowSelectionChange** (_Sel_)

_expression_ A variable that represents an [Application](Word.Application.md) object that has been declared with events in a class module. For more information about using events with the **Application** object, see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md).


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Sel_|Required| **Selection**|The text selected. If no text is selected, the Sel parameter returns either nothing or the first character to the right of the insertion point.|

## Example

This example applies bold formatting to the new selection. This code must be placed in a class module, and an instance of the class must be correctly initialized to see this example work; see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md) for directions on how to accomplish this.


```vb
Public WithEvents appWord as Word.Application 
 
Private Sub appWord_WindowSelectionChange _ 
 (ByVal Sel As Selection) 
 Sel.Font.Bold = True 
End Sub
```


## See also

- [Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]