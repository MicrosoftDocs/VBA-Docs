---
title: MailMerge object (Word)
keywords: vbawd10.chm2336
f1_keywords:
- vbawd10.chm2336
ms.prod: word
api_name:
- Word.MailMerge
ms.assetid: b228c4d6-9ca7-8795-12f6-d32e62844a83
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMerge object (Word)

Represents the mail merge functionality in Word.


## Remarks

Use the  **MailMerge** property to return the **MailMerge** object. The **MailMerge** object is always available regardless of whether the mail merge operation has begun. Use the **State** property to determine the status of the mail merge operation. The following example executes a mail merge if the active document is a main document with an attached data source.


```vb
If ActiveDocument.MailMerge.State = wdMainAndDataSource Then 
 ActiveDocument.MailMerge.Execute 
End If
```

The following example merges the main document with the first three records in the attached data source and then sends the results to the printer.




```vb
Set myMerge = ActiveDocument.MailMerge 
If myMerge.State = wdMainAndSourceAndHeader Or _ 
 myMerge.State = wdMainAndDataSource Then 
 With myMerge.DataSource 
 .FirstRecord = 1 
 .LastRecord = 3 
 End With 
End If 
With myMerge 
 .Destination = wdSendToPrinter 
 .Execute 
End With
```


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
