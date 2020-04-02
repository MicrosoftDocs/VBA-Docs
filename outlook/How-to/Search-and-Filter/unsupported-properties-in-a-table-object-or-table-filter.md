---
title: Unsupported Properties in a Table Object or Table Filter
ms.prod: outlook
ms.assetid: 0e37f03f-7677-ca29-d0b2-8b45c026e5f1
ms.date: 06/08/2019
localization_priority: Normal
---


# Unsupported Properties in a Table Object or Table Filter

This topic lists the properties that you cannot add to a **[Table](../../../api/Outlook.Table.md)** or use in a **Table** filter. You cannot add these properties through **[Columns.Add](../../../api/Outlook.Columns.Add.md)**, and you cannot specify these properties in a filter used by the following methods:


- **[Folder.GetTable](../../../api/Outlook.Folder.GetTable.md)**
    
- **[Search.GetTable](../../../api/Outlook.Search.GetTable.md)** (Note that the filter is derived from the **[Search](../../../api/Outlook.Search.md)** object returned by **[Application.AdvancedSearch](../../../api/Outlook.Application.AdvancedSearch.md)**)
    
- **[Table.FindRow](../../../api/Outlook.Table.FindRow.md)**
    
- **[Table.Restrict](../../../api/Outlook.Table.Restrict.md)**
    

| **Properties**| **In Table Object**| **In Table Filter**| **Comments**|
|:-----|:-----|:-----|:-----|
|Binary properties|Supported |Not supported|If you add a binary property to a **Table** referencing its namespace, the value of the property in the **Table** is in binary. You can use **[Row.BinaryToString](../../../api/Outlook.Row.BinaryToString.md)** to convert the value to a string.|
|Body properties, including **Body**, **HTMLBody**, **https://schemas.microsoft.com/mapi/proptag/0x10130102**<br> (for **PidTagHtml**), and **https://schemas.microsoft.com/mapi/proptag/0x10090102** (for **PidTagRtfCompressed**)|The **Body** property is supported with a condition that only the first 255 bytes of the value are stored in a **Table**. Other properties representing the body content in HTML or RTF are not supported. <br> Because only the first 255 bytes of **Body** is stored in a **Table**, if you want to obtain the full body content of an item in text or HTML, use the item's **EntryID** in **[GetItemFromID](../../../api/Outlook.NameSpace.GetItemFromID.md)** to obtain the item object. Then retrieve the full value of **Body** through the item object.|Only the **Body** property represented in text is supported in a filter. This means that the property must be referenced in a DASL filter as **urn:schemas:httpmail:textdescription**, and you cannot filter on any HTML tags in the body. To improve performance, use context indexer keywords in the filter to match strings in the body.||
|Computed properties, such as **AutoResolvedWinner** and **BodyFormat**. See below for a complete list of computed properties.|Not supported|Not supported|To obtain the value of a computed property for an item in a **Table**, use the item's **EntryID** in **GetItemFromID** to obtain the item object. Then retrieve the property value through the item object.|
|Multi-valued properties, such as **Categories**, **[Children](../../../api/Outlook.ContactItem.Children.md)**, **[Companies](../../../api/Outlook.ContactItem.Companies.md)**, and **[VotingOptions](../../../api/Outlook.MailItem.VotingOptions.md)**|Supported|Although both Jet and DASL filters both support multi-valued properties, use content indexing in DASL filters for more efficient filtering. For more information, see [Filtering Items Using a Comparison with a Keywords Property](filtering-items-using-a-comparison-with-a-keywords-property.md).|The format of the values of a multi-valued property in a **Table** depends on whether the property was added with its explicit built-in name or with a name referencing its namespace. If the property is added with its explicit built-in name, the value in the **Table** is a comma-delimited string. Otherwise, the value is a variant array. For more information, see [How to: Access the Values of a Multi-valued Property in a Table](access-the-values-of-a-multi-valued-property-in-a-table.md).|
|Properties returning an object, such as **Attachments**, **Parent**, **Recipients**, **RecurrencePattern**, and **UserProperties**.|Not supported if property is referenced by its explicit built-in name; supported if property is referenced by its namespace.|Not supported if property is expressed in a Jet query; supported if property is expressed in a DASL query.||


## Unsupported Computed Properties

If you attempt to add one of the computed properties listed below using **Columns.Add**, referencing the property either by the explicit property name or by namespace, you will get the error, **IDS_ERR_BLOCKED_PROPERTY**. To determine the value of these properties, obtain the item object using its Entry ID and then use the item object to determine the property value (as in  `object.property`):


- **AutoResolvedWinner**
    
- **BodyFormat**
    
- **Class**
    
- **ContactNames**
    
- **Companies**
    
- **[DLName](../../../api/Outlook.DistListItem.DLName.md)**
    
- **DownloadState**
    
- **FlagIcon**
    
- **HtmlBody**
    
- **InternetCodePage**
    
- **IsConflict**
    
- **IsMarkedAsTask**
    
- **MeetingWorkspaceURL**
    
- **MemberCount**
    
- **[Permission](../../../api/Outlook.MailItem.Permission.md)**
    
- **[PermissionService](../../../api/Outlook.MailItem.PermissionService.md)**
    
- **[RecurrenceState](../../../api/Outlook.AppointmentItem.RecurrenceState.md)**
    
- **[ResponseState](../../../api/Outlook.TaskItem.ResponseState.md)**
    
- **Saved**
    
- **Sent**
    
- **Submitted**
    
- **TaskSubject**
    
- **Unread**
    
- **[VotingOptions](../../../api/Outlook.MailItem.VotingOptions.md)**
    


If you attempt to use one of the computed properties listed below in a Jet filter (referencing the property by its explicit property name) for **Table.Restrict**, you will get the error, **IDS_ERR_ES_INVALIDRESTRICTION**: 


- **AutoResolvedWinner**
    
- **Body**
    
- **BodyFormat**
    
- **Class**
    
- **ContactNames**
    
- **Companies**
    
- **[CompanyLastFirstNoSpace](../../../api/Outlook.ContactItem.CompanyLastFirstNoSpace.md)**
    
- **[CompanyLastFirstSpaceOnly](../../../api/Outlook.ContactItem.CompanyLastFirstSpaceOnly.md)**
    
- **ContactNames**
    
- **[Contents](../../../api/Outlook.OutlookBarPane.Contents.md)**
    
- **ConversationIndex**
    
- **[DLName](../../../api/Outlook.DistListItem.DLName.md)**
    
- **DownloadState**
    
- **[Email1EntryID](../../../api/Outlook.ContactItem.Email1EntryID.md)**
    
- **[Email2EntryID](../../../api/Outlook.ContactItem.Email2EntryID.md)**
    
- **[Email3EntryID](../../../api/Outlook.ContactItem.Email3EntryID.md)**
    
- **EntryID**
    
- **HtmlBody**
    
- **InternetCodePage**
    
- **IsConflict**
    
- **IsMarkedAsTask**
    
- **[LastFirstAndSuffix](../../../api/Outlook.ContactItem.LastFirstAndSuffix.md)**
    
- **[LastFirstNoSpace](../../../api/Outlook.ContactItem.LastFirstNoSpace.md)**
    
- **[LastFirstNoSpaceAndSuffix](../../../api/Outlook.ContactItem.LastFirstNoSpaceAndSuffix.md)**
    
- **[LastFirstNoSpaceCompany](../../../api/Outlook.ContactItem.LastFirstNoSpaceCompany.md)**
    
- **[LastFirstSpaceOnly](../../../api/Outlook.ContactItem.LastFirstSpaceOnly.md)**
    
- **[LastFirstSpaceOnlyCompany](../../../api/Outlook.ContactItem.LastFirstSpaceOnlyCompany.md)**
    
- **MeetingWorkspaceURL**
    
- **MemberCount**
    
- **[NetMeetingAlias](../../../api/Outlook.ContactItem.NetMeetingAlias.md)**
    
- **NetMeetingServer**
    
- **[Permission](../../../api/Outlook.MailItem.Permission.md)**
    
- **[PermissionService](../../../api/Outlook.MailItem.PermissionService.md)**
    
- **[RecurrenceState](../../../api/Outlook.AppointmentItem.RecurrenceState.md)**
    
- **[ReceivedByEntryID](../../../api/Outlook.MailItem.ReceivedByEntryID.md)**
    
- **[ReceivedOnBehalfOfEntryID](../../../api/Outlook.MailItem.ReceivedOnBehalfOfEntryID.md)**
    
- **ReplyRecipients**
    
- **[ResponseState](../../../api/Outlook.TaskItem.ResponseState.md)**
    
- **Saved**
    
- **Sent**
    
- **Submitted**
    
- **TaskSubject**
    
- **[VotingOptions](../../../api/Outlook.MailItem.VotingOptions.md)**
    

 **Note** For a computed property such as **TaskSubject** or **IsMarkedAsTask**, you cannot add it to a **Table** using **Columns.Add** or filter it using **Table.Restrict**, if you reference the property with the explicit property name. However, you can add or filter on the property if you reference it by namespace, as in the following code sample in Visual Basic for Applications: 



```vb
Sub TableForIsMarkedAsTask() 
    Dim oT As Outlook.Table 
    Dim oRow As Outlook.Row 
    Dim filter As String 
    '0x0E2B0003 represents IsMarkedAsTask 
    filter = "@SQL=" & Chr(34) _ 
    & "https://schemas.microsoft.com/mapi/proptag/0x0E2B0003" & Chr(34) & " = 1" 
    'Table only contains rows for items where IsMarkedAsTask is True 
    Set oT = Application.Session.GetDefaultFolder(olFolderInbox).GetTable(filter) 
    oT.Columns.Add ("TaskStartDate") 
    oT.Columns.Add ("TaskDueDate") 
    oT.Columns.Add ("TaskCompletedDate") 
    'Use GUID/ID to represent TaskSubject 
    oT.Columns.Add ( _ 
        "https://schemas.microsoft.com/mapi/id/{00062008-0000-0000-C000-000000000046}/85A4001E") 
    Do Until oT.EndOfTable 
        Set oRow = oT.GetNextRow 
        Debug.Print oRow( _ 
        "https://schemas.microsoft.com/mapi/id/{00062008-0000-0000-C000-000000000046}/85A4001E"), _ 
        oRow("TaskStartDate"), oRow("TaskDueDate"), oRow("TaskCompletedDate") 
    Loop 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]