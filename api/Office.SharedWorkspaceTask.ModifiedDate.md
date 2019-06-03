---
title: SharedWorkspaceTask.ModifiedDate property (Office)
keywords: vbaof11.chm264010
f1_keywords:
- vbaof11.chm264010
ms.prod: office
api_name:
- Office.SharedWorkspaceTask.ModifiedDate
ms.assetid: 26b96d4d-b3ee-a9cc-2a00-73457820b3e1
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspaceTask.ModifiedDate property (Office)

Gets the date and time when the **SharedWorkspaceTask** object was last modified. Read-only.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_.**ModifiedDate**

_expression_ A variable that represents a **[SharedWorkspaceTask](Office.SharedWorkspaceTask.md)** object.


## Return value

Variant


## Example

The following example returns a list of shared workspace files whose date and time last modified is earlier than today.


```vb
Dim swsFile As Office.SharedWorkspaceFile 
    Dim dtmMidnight As Date 
    Dim dtmFileDate As Date 
    Dim strOlderFiles As String 
    dtmMidnight = CDate(FormatDateTime(Now, vbShortDate) & " 12:00:00 am") 
    For Each swsFile In ActiveWorkbook.SharedWorkspace.Files 
        dtmFileDate = swsFile.ModifiedDate 
        If dtmFileDate < dtmMidnight Then 
            strOlderFiles = strOlderFiles & swsFile.URL & vbCrLf 
        End If 
    Next 
    MsgBox "Files not modified today: " & vbCrLf & strOlderFiles, _ 
        vbInformation + vbOKOnly, "Older Files" 
    Set swsFile = Nothing
```


## See also

- [SharedWorkspaceTask object members](overview/Library-Reference/sharedworkspacetask-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]