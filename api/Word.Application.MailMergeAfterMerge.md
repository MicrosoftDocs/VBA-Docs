---
title: Application.MailMergeAfterMerge event (Word)
keywords: vbawd10.chm4000016
f1_keywords:
- vbawd10.chm4000016
ms.prod: word
api_name:
- Word.Application.MailMergeAfterMerge
ms.assetid: 6eed8afa-efe6-0eba-6ab8-6c3ffc4e812d
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.MailMergeAfterMerge event (Word)

Occurs after all records in a mail merge have merged successfully.


## Syntax

_expression_.**MailMergeAfterMerge** (_Doc_, _DocResult_)

_expression_ A variable that represents an **[Application](Word.Application.md)** object. 

For information about using events with the **Application** object, see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md).


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The mail merge main document.|
| _DocResult_|Required| **Document**|The document created from the mail merge|

## Example

This example displays a message stating that all records in the specified document are finished merging. If the document has been merged to a second document, the message includes the name of the new document. This example assumes that you have declared an application variable called MailMergeApp in your general declarations and have set the variable equal to the Word Application object.


```vb
Private Sub MailMergeApp_MailMergeAfterMerge(ByVal Doc As Document, _ 
 ByVal DocResult As Document) 
 If DocResult Is Nothing Then 
 MsgBox "Your mail merge on " & _ 
 Doc.Name & " is now finished." 
 
 Else 
 MsgBox "Your mail merge on " & _ 
 Doc.Name & " is now finished and " & _ 
 DocResult.Name & " has been created." 
 End If 
End Sub
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]