---
title: SharingItem.Type property (Outlook)
keywords: vbaol11.chm701
f1_keywords:
- vbaol11.chm701
ms.prod: outlook
api_name:
- Outlook.SharingItem.Type
ms.assetid: 1077b74f-38ee-8932-792d-64033bc66525
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.Type property (Outlook)

Returns or sets an **[OlSharingMsgType](Outlook.OlSharingMsgType.md)** constant that indicates the type of sharing message represented by the **[SharingItem](Outlook.SharingItem.md)**. Read/write.


## Syntax

_expression_.**Type**

 _expression_ An expression that returns a [SharingItem](Outlook.SharingItem.md) object.


## Remarks

An error occurs if you attempt to set this property after the sharing message has been sent or received, or if you attempt to set this property to  **olSharingMsgTypeResponseAllow** or **olSharingMsgTypeResponseDeny**.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]