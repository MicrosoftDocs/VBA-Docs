---
title: PpRemoveDocInfoType enumeration (PowerPoint)
ms.prod: powerpoint
api_name:
- PowerPoint.PpRemoveDocInfoType
ms.assetid: 76cb213a-34a4-8b5e-6e9d-9fc7528c7574
ms.date: 12/10/2018
localization_priority: Normal
---


# PpRemoveDocInfoType enumeration (PowerPoint)

Constants that specify the type of document information, passed to the **[RemoveDocumentInformation](powerpoint.presentation.removedocumentinformation.md)** method of the **[Presentation](powerpoint.presentation.md)** object.

|Name|Value|Description|
|:---|:----|:----------|
|**ppRDIAll**|99|Remove all document information.|
|**ppRDIAtMentions**|18|Remove resolved @mentioned users from comments.|
|**ppRDIComments**|1|Remove comments.|
|**ppRDIContentType**|16|Remove content type information.|
|**ppRDIDocumentManagementPolicy**|15|Remove document management policy information.|
|**ppRDIDocumentProperties**|8|Remove document properties.|
|**ppRDIDocumentServerProperties**|14|Remove document server properties.|
|**ppRDIDocumentWorkspace**|10|Remove document workspace information. |
|**ppRDIInkAnnotations**|11|Remove Ink annotations.<br/><br/>**NOTE**: This constant has been deprecated, but it remains part of the object model for backward compatibility. You should not use it in new applications.|
|**ppRDIPublishPath**|13|Remove publication path information.|
|**ppRDIRemovePersonalInformation**|4|Remove personal information.|
|**ppRDISlideUpdateInformation**|17|Remove slide update information.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]