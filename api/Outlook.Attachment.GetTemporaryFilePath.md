---
title: Attachment.GetTemporaryFilePath method (Outlook)
keywords: vbaol11.chm3522
f1_keywords:
- vbaol11.chm3522
ms.prod: outlook
api_name:
- Outlook.Attachment.GetTemporaryFilePath
ms.assetid: 3313582b-6241-7a59-0c03-b8af36a17d3d
ms.date: 06/08/2017
localization_priority: Normal
---


# Attachment.GetTemporaryFilePath method (Outlook)

Returns the full path to the attached file that is in a temporary files folder. Read-only.


## Syntax

_expression_. `GetTemporaryFilePath`

_expression_ A variable that represents an '[Attachment](Outlook.Attachment.md)' object.


## Return value

Returns a  **String** that represents the full path to the temporary attachment file.


## Remarks

The  **GetTemporaryFilePath** method is only valid for those attachments whose **[Type](Outlook.Attachment.Type.md)** property is **OlAttachmentType.olByValue**. That means that the attachment is a copy and that the copy can be accessed even if the original file is removed. For other attachment types, the **GetTemporaryFilePath** method returns an error.

 **GetTemporaryFilePath** also returns an error when accessing an **[Attachment](Outlook.Attachment.md)** object in an **[Attachments](Outlook.Attachments.md)** collection or in the **[AttachmentSelection](Outlook.AttachmentSelection.md)** object. Use **GetTemporaryFilePath** only in attachment event callbacks listed below for various Microsoft Outlook items:


-  **AttachmentAdd**
    
-  **AttachmentRead**
    
-  **AttachmentRemove**
    
-  **BeforeAttachmentAdd**
    
-  **BeforeAttachmentPreview**
    
-  **BeforeAttachmentRead**
    
-  **BeforeAttachmentSave**
    
-  **BeforeAttachmentWriteToTempFile**
    



## See also


[Attachment Object](Outlook.Attachment.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]