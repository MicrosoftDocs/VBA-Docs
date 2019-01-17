---
title: SharedWorkspaceTask.Status property (Office)
keywords: vbaof11.chm264003
f1_keywords:
- vbaof11.chm264003
ms.prod: office
api_name:
- Office.SharedWorkspaceTask.Status
ms.assetid: de1e6222-67cb-107d-ad59-7d3ea38d5283
ms.date: 06/08/2017
localization_priority: Normal
---


# SharedWorkspaceTask.Status property (Office)

Gets or sets the status of the specified shared workspace task. Read/write .

> [!NOTE] 
> Beginning with Microsoft Office 2010, this object or member has been deprecated and should not be used.


## Syntax

_expression_. `Status`

 _expression_ Required. A variable that represents a '[SharedWorkspaceTask](Office.SharedWorkspaceTask.md)' object.


## Remarks

The shared workspace task schema on the server can be customized. Customization of the schema may affect the task status enumeration when the  **Add** or **Save** method is called. **Status** property values are mapped as follows:




- Downloaded values 1 through 5 are mapped to  **msoSharedWorkspaceTaskStatus** enumeration values 1 through 5. Schema values beyond 5 are mapped to enumeration value 1 (**msoSharedWorkspaceTaskStatusInProgress**).
    
- Uploaded enumeration values 1 through 5 are mapped to schema values 1 through 5. If a user-specified value does not map to any value defined in the schema, the user-specified value is silently ignored and the  **Status** property is not updated on the server.
    



## Example

The following example displays a list of all tasks in the current shared workspace whose status is not set to Complete.


```vb
    Dim swsTask As Office.SharedWorkspaceTask 
    Dim strTaskStatus As String 
    For Each swsTask In ActiveWorkbook.SharedWorkspace.Tasks 
        If swsTask.Status <> msoSharedWorkspaceTaskStatusCompleted Then 
            strTaskStatus = strTaskStatus &amp; swsTask.Title &amp; vbCrLf 
        End If 
    Next 
    MsgBox "The following tasks have not been completed:" &amp; vbCrLf &amp; _ 
        strTaskStatus, vbInformation + vbOKOnly, "Incomplete Tasks" 
    Set swsTask = Nothing 

```


## See also


[SharedWorkspaceTask Object](Office.SharedWorkspaceTask.md)



[SharedWorkspaceTask Object Members](./overview/Library-Reference/sharedworkspacetask-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]