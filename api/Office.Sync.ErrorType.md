---
title: Sync.ErrorType Property (Office)
keywords: vbaof11.chm277005
f1_keywords:
- vbaof11.chm277005
ms.prod: office
api_name:
- Office.Sync.ErrorType
ms.assetid: 6663e5f6-b90e-29f8-2ff9-f9fb8bda76f0
ms.date: 06/08/2017
---


# Sync.ErrorType Property (Office)

Gets a  **MsoSyncErrorType** constant which indicates the type of the most recent document synchronization error. Read-only.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

 _expression_. `ErrorType`

 _expression_ A variable that returns a '[Sync](Office.Sync.md)' object.


### Return value

MsoSyncErrorType


## Remarks

Use the  **ErrorType** property to determine the type of the most recent document synchronization error. Not all document synchronization problems raise trappable run-time errors. After performing an operation using the **Sync** object, it's a good idea to check the **Status** property; if the **Status** property is **msoSyncStatusError**, check the **ErrorType** property for additional information on the type of error that has occurred.


## See also


[Sync Object](Office.Sync.md)



[Sync Object Members](./overview/Library-Reference/sync-members-office.md)

