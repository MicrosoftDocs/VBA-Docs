---
title: Application.MailMergeRecipientListClose event (Publisher)
keywords: vbapb10.chm268435488
f1_keywords:
- vbapb10.chm268435488
ms.prod: publisher
api_name:
- Publisher.Application.MailMergeRecipientListClose
ms.assetid: 4fb77771-9897-8623-f4e7-61f631f04922
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.MailMergeRecipientListClose event (Publisher)

Fires when the user closes the **Mail Merge Recipients** dialog box (from the **Mail Merge** or **Email Merge** task pane, choose **Edit Recipient List**). Also fires when the user closes the **Catalog Merge Product List** dialog box, which opens when the user chooses **Edit Product List** in the **Catalog Merge** task pane.


## Syntax

_expression_.**MailMergeRecipientListClose** (_Doc_)

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Doc_|Required| **Document**|The current publication.|

## Remarks

For more information about using events with the **Application** object, see [Using events with the Application object](../publisher/Concepts/using-events-with-the-application-object-publisher.md).


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to handle the **MailMergeRecipientListClose** event. It displays a message notifying the user that the string described earlier was displayed.

```vb
Private Sub pubApplication_MailMergeRecipientListClose(ByVal Doc As Document) 
 MsgBox "The Mail Merge Recipients dialog box has closed." 
End Sub
```

<br/>

For this event to occur, you must place the following line of code in the General Declarations section of your module.

```vb
Private WithEvents pubApplication As Application
```

<br/>

You then must run the following initialization procedure.

```vb
Public Sub Initialize_pubApplication() 
 Set pubApplication = Publisher.Application 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]