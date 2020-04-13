---
title: Application.DocumentBeforePrint event (Word)
keywords: vbawd10.chm400006
f1_keywords:
- vbawd10.chm400006
ms.prod: word
api_name:
- Word.Application.DocumentBeforePrint
ms.assetid: 0736197a-5770-7e00-9882-86be0579c83e
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.DocumentBeforePrint event (Word)

Occurs before any open document is printed.


## Syntax

_expression_.**DocumentBeforePrint** (_Doc As Document_**, **_Cancel As Boolean_**)

_expression_ A variable that represents an '[Application](Word.Application.md)' object declared with events in a class module.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The document that's being printed.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the document isn't printed when the procedure is finished.|

## Remarks

For more information about using events with the **Application** object, see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md).


## Example

This example prompts the user for a yes or no response before printing any document. This code must be placed in a class module, and an instance of the class must be correctly initialized to see this example work; see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md)for directions on how to accomplish this.


```vb
Public WithEvents appWord as Word.Application 
 
Private Sub appWord_DocumentBeforePrint _ 
 (ByVal Doc As Document, _ 
 Cancel As Boolean) 
 
 Dim intResponse As Integer 
 
 intResponse = MsgBox("Have you checked the " _ 
 & "printer for letterhead?", _ 
 vbYesNo) 
 
 If intResponse = vbNo Then Cancel = True 
End Sub
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]