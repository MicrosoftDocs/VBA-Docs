---
title: SharedWorkspace.LastRefreshed property (Office)
keywords: vbaof11.chm276013
f1_keywords:
- vbaof11.chm276013
ms.prod: office
api_name:
- Office.SharedWorkspace.LastRefreshed
ms.assetid: 426c53dd-3f3a-c638-2559-c02f62f374ff
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspace.LastRefreshed property (Office)

Gets the date and time when the **Refresh** method was most recently called. Read-only.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_.**LastRefreshed**

_expression_ A variable that represents a **[SharedWorkspace](Office.SharedWorkspace.md)** object.


## Remarks

The **LastRefreshed** property raises an error if the **[Refresh](Office.SharedWorkspace.Refresh.md)** method has never been called.


## Example

The following example refreshes the shared workspace if it has not been refreshed in the last 3 minutes. The example also handles the error condition where the workspace has not yet been refreshed.


```vb
    On Error GoTo err_NeverRefreshed 
    If DateDiff("s", ActiveWorkbook.SharedWorkspace.LastRefreshed, Now) > 180 Then 
        ActiveWorkbook.SharedWorkspace.Refresh 
    End If 
    Exit Sub 
     
err_NeverRefreshed: 
          ActiveWorkbook.SharedWorkspace.Refresh 

```


## See also

- [SharedWorkspace object members](overview/Library-Reference/sharedworkspace-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]