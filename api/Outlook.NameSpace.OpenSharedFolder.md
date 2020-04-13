---
title: NameSpace.OpenSharedFolder method (Outlook)
keywords: vbaol11.chm788
f1_keywords:
- vbaol11.chm788
ms.prod: outlook
api_name:
- Outlook.NameSpace.OpenSharedFolder
ms.assetid: 907efeab-8a37-98a6-f241-0a051f03f472
ms.date: 06/08/2017
localization_priority: Normal
---


# NameSpace.OpenSharedFolder method (Outlook)

Opens a shared folder referenced through a URL or file name.


## Syntax

_expression_. `OpenSharedFolder`( `_Path_` , `_Name_` , `_DownloadAttachments_` , `_UseTTL_` )

 _expression_ An expression that returns a [NameSpace](Outlook.NameSpace.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Path_|Required| **String**|The URL or local file name of the shared folder to be opened.|
| _Name_|Optional| **Variant**|The name of the Really Simple Syndication (RSS) feed or Webcal calendar. This parameter is ignored for other shared folder types.|
| _DownloadAttachments_|Optional| **Variant**|Indicates whether to download enclosures (for RSS feeds) or attachments (for Webcal calendars.) This parameter is ignored for other shared folder types.|
| _UseTTL_|Optional| **Variant**|Indicates whether the Time To Live (TTL) setting in an RSS feed or WebCal calendar should be used. This parameter is ignored for other shared folder types.|

## Return value

A **[Folder](Outlook.Folder.md)** object that represents the shared folder.


## Remarks

This method is used to access the following shared folder types:


- Webcal calendars (webcal:// _mysite_ / _mycalendar_ )
    
- RSS feeds (feed:// _mysite_ / _myfeed_ )
    
- Microsoft SharePoint Foundation folders (stssync:// _mysite_ / _myfolder_ )
    
- iCalendar calendar (.ics) files
    
- vCard contact (.vcf) files
    
- Outlook message (.msg) files
    

> [!NOTE] 
> This method does not support iCalendar appointment (.ics) files. To open iCalendar appointment files, you can use the  **[OpenSharedItem](Outlook.NameSpace.OpenSharedItem.md)** method of the **[NameSpace](Outlook.NameSpace.md)** object.

You can use the  **[GetSharedDefaultFolder](Outlook.NameSpace.GetSharedDefaultFolder.md)** method of the **Namespace** object to share default folders, such as the Inbox folder, in Exchange.


## Example

The following Visual Basic for Applications (VBA) example opens and displays a Webcal calendar. 


```vb
Public Sub OpenSharedHolidayCalendar() 
 
 
 
 Dim oNamespace As NameSpace 
 
 Dim oFolder As Folder 
 
 
 
 On Error GoTo ErrRoutine 
 
 
 
 Set oNamespace = Application.GetNamespace("MAPI") 
 
 Set oFolder = oNamespace.OpenSharedFolder( _ 
 
 "webcal://icalx.com/public/icalshare/US32Holidays.ics") 
 
 oFolder.Display 
 
 
 
EndRoutine: 
 
 On Error GoTo 0 
 
 Set oFolder = Nothing 
 
 Set oNamespace = Nothing 
 
Exit Sub 
 
 
 
ErrRoutine: 
 
 MsgBox Err.Description, vbOKOnly, Err.Number & " - " & Err.Source 
 
 GoTo EndRoutine 
 
End Sub
```


## See also


[NameSpace Object](Outlook.NameSpace.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]