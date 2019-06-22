---
title: Application.MailMergeWizardSendToCustom event (Word)
keywords: vbawd10.chm4000022
f1_keywords:
- vbawd10.chm4000022
ms.prod: word
api_name:
- Word.Application.MailMergeWizardSendToCustom
ms.assetid: b5dcd912-f1b5-96d6-3221-d294211b6611
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.MailMergeWizardSendToCustom event (Word)

Occurs when the custom button is clicked during step six of the Mail Merge Wizard.


## Syntax

_expression_.**MailMergeWizardSendToCustom** (_Doc As Document_**)

_expression_ A variable that represents an '[Application](Word.Application.md)' object that has been declared with events in a class module.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The mail merge main document.|

## Remarks

Use the  **ShowSendToCustom** property to create a custom button on the sixth step of the Mail Merge Wizard.

For information about using events with the  **Application** object, see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md).


## Example

This example executes a merge to a fax machine when a user clicks the custom destination button. This example assumes that the user has access to a custom destination button, fax numbers are included for each record in the data source, and an application variable called MailMergeApp has been declared and set equal to the Word Application object.


```vb
Private Sub MailMergeApp_MailMergeWizardSendToCustom(ByVal Doc As Document) 
 With Doc.MailMerge 
 .Destination = wdSendToFax 
 .Execute 
 End With 
End Sub
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]