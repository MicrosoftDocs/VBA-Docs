---
title: Application.MailMergeWizardFollowUpCustom event (Publisher)
keywords: vbapb10.chm268435490
f1_keywords:
- vbapb10.chm268435490
ms.prod: publisher
api_name:
- Publisher.Application.MailMergeWizardFollowUpCustom
ms.assetid: ac8cb695-69a4-83f7-8e13-66762f52f611
ms.date: 06/05/2019
localization_priority: Normal
---


# Application.MailMergeWizardFollowUpCustom event (Publisher)

Fires when the string that appears as the fourth item under **Prepare to follow-up on this mailing** in the third **Mail Merge** task pane in the Microsoft Publisher user interface is chosen.


## Syntax

_expression_.**MailMergeWizardFollowUpCustom** (_Doc_)

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Doc_|Required| **Document**|The current publication.|

## Remarks

You can use the **[ShowFollowUpCustom](Publisher.Application.ShowFollowUpCustom.md)** property to display this string.

For more information about using events with the **Application** object, see [Using events with the Application object](../publisher/Concepts/using-events-with-the-application-object-publisher.md).


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to handle the **MailMergeWizardFollowUpCustom** event. It displays a message notifying the user that the string described earlier was displayed.

```vb
Private Sub pubApplication_MailMergeWizardFollowUpCustom(ByVal Doc As Document) 
 MsgBox "The FollowUpCustom string is clicked." 
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