---
title: MailMerge.ViewMailMergeFieldCodes property (Word)
keywords: vbawd10.chm153092102
f1_keywords:
- vbawd10.chm153092102
ms.prod: word
api_name:
- Word.MailMerge.ViewMailMergeFieldCodes
ms.assetid: f39e93d8-bc80-8a3d-8bfc-5d6fbb0162f4
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMerge.ViewMailMergeFieldCodes property (Word)

 **True** if merge field names are displayed in a mail merge main document. **False** if information from the current record is displayed. Read/write **Long**.


## Syntax

_expression_. `ViewMailMergeFieldCodes`

 _expression_ An expression that returns a '[MailMerge](Word.MailMerge.md)' object.


## Remarks

If the active document isn't a mail merge main document, this property causes an error. To view merge field names or their results, set the  **[ShowFieldCodes](Word.View.ShowFieldCodes.md)** property to **False**.


## Example

This example displays the mail merge fields in Main.doc.


```vb
ActiveDocument.ActiveWindow.View.ShowFieldCodes = False 
With Documents("Main.doc") 
 .Activate 
 .MailMerge.ViewMailMergeFieldCodes = True 
End With
```

If the active document is set up for a mail merge operation, this example displays the current record information in the main document.




```vb
ActiveDocument.ActiveWindow.View.ShowFieldCodes = False 
Set myMerge = ActiveDocument.MailMerge 
If myMerge.State = wdMainAndSourceAndHeader Or _ 
 myMerge.State = wdMainAndDataSource Then 
 myMerge.ViewMailMergeFieldCodes = False 
End If
```


## See also


[MailMerge Object](Word.MailMerge.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]