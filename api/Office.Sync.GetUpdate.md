---
title: Sync.GetUpdate method (Office)
keywords: vbaof11.chm277007
f1_keywords:
- vbaof11.chm277007
ms.prod: office
api_name:
- Office.GetUpdate
ms.assetid: a92c0096-fcf2-2754-31e6-2b20a5841463
ms.date: 01/25/2019
localization_priority: Normal
---


# Sync.GetUpdate method (Office)

Compares the local version of the shared document to the version on the server.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_.**GetUpdate**

_expression_ A variable that represents a **[Sync](Office.Sync.md)** object.


## Remarks

Use the **GetUpdate** method to compare the local version of the shared document to the version on the server and to refresh the sync status.

Not all document synchronization problems raise trappable run-time errors. After performing an operation by using the **Sync** object, it's a good idea to check the **Status** property; if the **Status** property is **[msoSyncStatusError](office.msosyncstatustype.md)**, check the **ErrorType** property for additional information about the type of error that has occurred.

In many circumstances, the best way to resolve an error condition is to call the **GetUpdate** method. For example, if a call to **PutUpdate** results in an error condition, a call to **GetUpdate** will reset the status to **msoSyncStatusLocalChanges**.


## Example

The following example compares the local and server copies of the document by using the **GetUpdate** method and reports whether the server has a newer copy.


```vb
    Dim objSync As Office.Sync 
    Dim strStatus As String 
    Set objSync = ActiveDocument.Sync 
    objSync.GetUpdate 
    If objSync.Status = msoSyncStatusNewerAvailable Then 
        strStatus = "A newer version is available on the server." 
        MsgBox strStatus, vbInformation + vbOKOnly, "Sync Information" 
    End If 
    Set objSync = Nothing 

```

## See also

- [Sync object members](overview/Library-Reference/sync-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]