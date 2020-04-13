---
title: SharingItem.OriginatorDeliveryReportRequested property (Outlook)
keywords: vbaol11.chm642
f1_keywords:
- vbaol11.chm642
ms.prod: outlook
api_name:
- Outlook.SharingItem.OriginatorDeliveryReportRequested
ms.assetid: 7864b388-fa76-14cd-1f1c-f9f2958ec1bb
ms.date: 06/08/2017
localization_priority: Normal
---


# SharingItem.OriginatorDeliveryReportRequested property (Outlook)

Returns or sets a **Boolean** value that determines whether the originator of the **[SharingItem](Outlook.SharingItem.md)** will receive a delivery report. Read/write.


## Syntax

_expression_. `OriginatorDeliveryReportRequested`

_expression_ A variable that represents a [SharingItem](Outlook.SharingItem.md) object.


## Remarks

Each transport provider that handles your message sends you a single delivery notification containing the names and addresses of each recipient to whom it was delivered. Delivery does not imply that the message has been read.  **True** if the originator requested a delivery receipt on the message.

The **OriginatorDeliveryReportRequested** property corresponds to the MAPI property **PidTagOriginatorDeliveryReportRequested**.


## See also


[SharingItem Object](Outlook.SharingItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]