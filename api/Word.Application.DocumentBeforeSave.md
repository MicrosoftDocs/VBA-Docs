---
title: Application.DocumentBeforeSave event (Word)
keywords: vbawd10.chm400007
f1_keywords:
- vbawd10.chm400007
ms.prod: word
api_name:
- Word.Application.DocumentBeforeSave
ms.assetid: cc1c6ec3-0e9e-5147-78a5-3a0c47fd5e90
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.DocumentBeforeSave event (Word)

Occurs before any open document is saved.


## Syntax

_expression_.**DocumentBeforeSave** (_DocAs Document_**, **_SaveAsUIAs Boolean_**, **_CancelAs Boolean_**)

_expression_ A variable that represents an '[Application](Word.Application.md)' object declared with events in a class module.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The document that is being saved.|
| _SaveAsUI_|Required| **Boolean**| **True** if the **Save As** dialog box is displayed, whether to save a new document, in response to the **Save** command; or in response to the **Save As** command; or in response to the **SaveAs** or **SaveAs2** method.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True**, the document is not saved when the procedure is finished.|

## Remarks

For more information about using events with the **Application** object, see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md).


## Example

This example prompts the user for a yes or no response before saving any document. This code must be placed in a class module, and an instance of the class must be correctly initialized to see this example work; see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md) for directions on how to accomplish this.


```vb
Public WithEvents appWord as Word.Application 
 
Private Sub appWord_DocumentBeforeSave _ 
 (ByVal Doc As Document, _ 
 SaveAsUI As Boolean, _ 
 Cancel As Boolean) 
 
 Dim intResponse As Integer 
 
 intResponse = MsgBox("Do you really want to " _ 
 & "save the document?", _ 
 vbYesNo) 
 
 If intResponse = vbNo Then Cancel = True 
End Sub
```


## See also


[Application Object](Word.Application.md)

[AutoSave](../Library-Reference/Concepts/how-autosave-impacts-addins-and-macros.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
