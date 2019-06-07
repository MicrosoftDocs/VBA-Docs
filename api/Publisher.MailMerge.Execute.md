---
title: MailMerge.Execute method (Publisher)
keywords: vbapb10.chm6225940
f1_keywords:
- vbapb10.chm6225940
ms.prod: publisher
api_name:
- Publisher.MailMerge.Execute
ms.assetid: edcabcc5-f2ce-53ce-d422-0d6fcb5f8a33
ms.date: 06/08/2019
localization_priority: Normal
---


# MailMerge.Execute method (Publisher)

Performs the specified mail merge or catalog merge operation. Returns a **[Document](Publisher.Document.md)** object that represents the new or existing publication specified as the destination of the merge results. Returns **Nothing** if the merge is executed to a printer.


## Syntax

_expression_.**Execute** (_Pause_, _Destination_, _FileName_)

_expression_ A variable that represents a **[MailMerge](Publisher.MailMerge.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Pause_|Required| **Boolean**| **True** to have Microsoft Publisher pause and display a troubleshooting dialog box if a merge error is found. **False** to ignore errors during a mail merge or catalog merge.|
|_Destination_|Optional| **[PbMailMergeDestination](publisher.pbmailmergedestination.md)**|The destination of the mail merge or catalog merge results. Can be one of the **PbMailMergeDestination** constants; the default is **pbSendToPrinter**. Specifying **pbSendToPrinter** for a catalog merge results in a run-time error.|
|_FileName_|Optional| **String**|The file name of the publication to which you want to append the catalog merge results.|

## Return value

Document


## Example

This example executes a mail merge if the active publication is a main document with an attached data source.

```vb
Sub ExecuteMerge() 
 Dim mrgDocument As MailMerge 
 Set mrgDocument = ActiveDocument.MailMerge 
 If mrgDocument.DataSource.ConnectString <> "" Then 
 mrgDocument.Execute Pause:=False 
 End If 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]