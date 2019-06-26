---
title: CustomLabel object (Word)
keywords: vbawd10.chm2325
f1_keywords:
- vbawd10.chm2325
ms.prod: word
api_name:
- Word.CustomLabel
ms.assetid: a89ff4e1-ff8a-8a8f-afa2-6071bb49355b
ms.date: 06/08/2017
localization_priority: Normal
---


# CustomLabel object (Word)

Represents a custom mailing label. The  **CustomLabel** object is a member of the **[CustomLabels](Word.customlabels.md)** collection. The **CustomLabels** collection contains all the custom mailing labels listed in the **Label Options** dialog box.


## Remarks

Use  **[CustomLabels](Word.MailingLabel.CustomLabels.md)** (Index), where Index is the custom label name or index number, to return a single **CustomLabel** object. The following example creates a new document with an existing custom label layout named "My Labels."


```vb
Set ML = Application.MailingLabel 
If ML.CustomLabels("My Labels").Valid = True Then 
 ML.CreateNewDocument Name:="My Labels" 
Else 
 MsgBox "The My Labels custom label is not available" 
End If
```

The index number represents the position of the custom mailing label in the  **[CustomLabels](Word.customlabels.md)** collection. The following example displays the name of the first custom mailing label.




```vb
If Application.MailingLabel.CustomLabels.Count >= 1 Then 
 MsgBox Application.MailingLabel.CustomLabels(1).Name 
End If
```


> [!NOTE] 
>  **CustomLabel** objects are sorted alphabetically in the **[CustomLabels](Word.customlabels.md)** collection and their index numbers are dynamically reassigned as the contents of the collection change. For that reason, it is safer to refer to a specific **CustomLabel** object by name rather than by index number.

Use the  **[Add](Word.CustomLabels.Add.md)** method to create a custom label. The following example adds a custom mailing label named "My Label" and sets the page size.




```vb
Set ML = _ 
 Application.MailingLabel.CustomLabels.Add(Name:="My Labels", _ 
 DotMatrix:=False) 
ML.PageSize = wdCustomLabelA4
```


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]