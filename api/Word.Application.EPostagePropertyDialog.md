---
title: Application.EPostagePropertyDialog event (Word)
keywords: vbawd10.chm4000014
f1_keywords:
- vbawd10.chm4000014
ms.prod: word
api_name:
- Word.Application.EPostagePropertyDialog
ms.assetid: 6d48fb9b-7826-2897-4deb-bde202fbd05b
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.EPostagePropertyDialog event (Word)

Occurs when a user clicks the **E-postage Properties** (**Labels and Envelopes** dialog box) button or **Print Electronic Postage** button.


## Syntax

_expression_.**EPostagePropertyDialog** (_Doc_)

_expression_ A variable that represents an '[Application](Word.Application.md)' object that has been declared with events in a class module. For information about using events with the **Application** object, see [Using events with the Application object](../word/Concepts/Objects-Properties-Methods/using-events-with-the-application-object-word.md).


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Doc_|Required| **Document**|The name of the document to which to add electronic postage.|

## Remarks

This event allows a third-party software application to intercept and show their properties dialog box.


## Example

This example displays a message when a user clicks either the **Add Electronic Postage** button or the **Print Electronic Postage** button.


```vb
Private Sub AppWord_EPostagePropertyDialog(ByVal Doc As Document) 
 MsgBox "You have clicked a button to " & _ 
 "display the ePostage property dialog box." 
End Sub
```


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]