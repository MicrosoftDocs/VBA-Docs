---
title: MailMerge Object (Publisher)
keywords: vbapb10.chm6291455
f1_keywords:
- vbapb10.chm6291455
ms.prod: publisher
api_name:
- Publisher.MailMerge
ms.assetid: 028e1e42-c61c-9b2b-4aec-d6a184504ec1
ms.date: 06/08/2017
---


# MailMerge Object (Publisher)

Represents the mail merge and catalog merge functionality in Microsoft Publisher.


## Example

Use the  **[MailMerge](./Publisher.Document.MailMerge.md)** property to return the **MailMerge** object. The **MailMerge** object is always available regardless of whether the mail merge or catalog merge operation has begun. The following example merges and prints the main publication with the first three records in the attached data source.


```vb
Sub SelectiveMerge() 
 Dim mrgMain As MailMerge 
 Set mrgMain = ActiveDocument.MailMerge 
 With mrgMain.DataSource 
 .FirstRecord = 1 
 .LastRecord = 3 
 End With 
 mrgMain.Execute True 
End Sub
```


## Methods



|Name|
|:-----|
|[CreateShortcut](./Publisher.MailMerge.CreateShortcut.md)|
|[Execute](./Publisher.MailMerge.Execute.md)|
|[ExportRecipientList](./Publisher.MailMerge.ExportRecipientList.md)|
|[OpenDataSource](./Publisher.MailMerge.OpenDataSource.md)|
|[ShowWizardEx](./Publisher.MailMerge.ShowWizardEx.md)|

## Properties



|Name|
|:-----|
|[Application](./Publisher.MailMerge.Application.md)|
|[DataSource](./Publisher.MailMerge.DataSource.md)|
|[DocumentUpdating](./Publisher.MailMerge.DocumentUpdating.md)|
|[EmailMergeEnvelope](./Publisher.MailMerge.EmailMergeEnvelope.md)|
|[Parent](./Publisher.MailMerge.Parent.md)|
|[SuppressBlankLines](./Publisher.MailMerge.SuppressBlankLines.md)|
|[Type](./Publisher.MailMerge.Type.md)|
|[ViewMailMergeFieldCodes](./Publisher.MailMerge.ViewMailMergeFieldCodes.md)|
|[WizardState](./Publisher.MailMerge.WizardState.md)|

