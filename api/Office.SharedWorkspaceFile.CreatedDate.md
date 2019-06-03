---
title: SharedWorkspaceFile.CreatedDate property (Office)
keywords: vbaof11.chm266003
f1_keywords:
- vbaof11.chm266003
ms.prod: office
api_name:
- Office.SharedWorkspaceFile.CreatedDate
ms.assetid: c3a45dbd-c6b2-3046-2388-ed23ca7e36f0
ms.date: 01/24/2019
localization_priority: Normal
---


# SharedWorkspaceFile.CreatedDate property (Office)

Gets the date and time when the shared workspace object was created. Read-only.

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.

## Syntax

_expression_.**CreatedDate**

_expression_ A variable that represents a **[SharedWorkspaceFile](Office.SharedWorkspaceFile.md)** object.


## Return value

Variant


## Example

The following example returns a list of shared workspace files whose date and time created is earlier than today.


```vb
 Dim swsFile As Office.SharedWorkspaceFile 
 Dim dtmMidnight As Date 
 Dim dtmFileDate As Date 
 Dim strOlderFiles As String 
 dtmMidnight = CDate(FormatDateTime(Now, vbShortDate) & " 12:00:00 am") 
 For Each swsFile In ActiveWorkbook.SharedWorkspace.Files 
 dtmFileDate = swsFile.CreatedDate 
 If dtmFileDate < dtmMidnight Then 
 strOlderFiles = strOlderFiles & swsFile.URL & vbCrLf 
 End If 
 Next 
 MsgBox "Files older than today: " & vbCrLf & strOlderFiles, _ 
 vbInformation + vbOKOnly, "Older Files" 
 Set swsFile = Nothing 
 

```




## See also

- [SharedWorkspaceFile object members](overview/Library-Reference/sharedworkspacefile-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]