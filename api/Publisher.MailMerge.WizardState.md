---
title: MailMerge.WizardState property (Publisher)
keywords: vbapb10.chm6225929
f1_keywords:
- vbapb10.chm6225929
ms.prod: publisher
api_name:
- Publisher.MailMerge.WizardState
ms.assetid: a237cb3f-2c03-5f62-fa67-d4aa7703389d
ms.date: 06/08/2019
localization_priority: Normal
---


# MailMerge.WizardState property (Publisher)

Returns or sets a **Long** indicating the current mail merge wizard step for a publication. The **WizardState** property returns a number that equates to the current mail merge wizard step; a zero (0) means that the mail merge wizard is closed. Read/write.


## Syntax

_expression_.**WizardState**

_expression_ A variable that represents a **[MailMerge](Publisher.MailMerge.md)** object.


## Return value

Long


## Example

This example displays the mail merge wizard if it is closed.

```vb
Sub ShowMergeWizard() 
 With ActiveDocument.MailMerge 
 If .WizardState = 0 Then 
 .ShowWizard 
 End If 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]