---
title: Category Object (Outlook)
keywords: vbaol11.chm3177
f1_keywords:
- vbaol11.chm3177
ms.prod: outlook
api_name:
- Outlook.Category
ms.assetid: 143ef095-54b0-cbe2-e356-632029061ac2
ms.date: 06/08/2017
---


# Category Object (Outlook)

Represents a user-defined category by which Outlook items can be grouped.


## Remarks

Microsoft Outlook provides a categorization system with which Outlook items can be easily identified and grouped into user-defined categories. The  **Category** object represents a user-defined category.

Use the  **[Add](Outlook.Categories.Add.md)** method of the **[Categories](Outlook.NameSpace.Categories.md)** property for the **[NameSpace](Outlook.NameSpace.md)** object to create a new **Category** object, adding the category to the Master Category List for that namespace.

Use the  **[Name](Outlook.Category.Name.md)** property to specify the name of the category, the **[Color](Outlook.Category.Color.md)** property to specify the color displayed for that category, and the **[ShortcutKey](Outlook.Category.ShortcutKey.md)** property to specify the shortcut key used to assign that category to an Outlook item in the Outlook user interface. Use the **[CategoryID](Outlook.Category.CategoryID.md)** property to retrieve the unique identifer for a category.


### Assigning Categories to Items

Categories can be assigned to Outlook items by specifying the names of the appropriate  **Category** objects in a comma-delimited string in the **Categories** property of the following objects:


|||
|:-----|:-----|
|**[AppointmentItem](appointmentitem-object-outlook.md)**|**[RemoteItem](remoteitem-object-outlook.md)**|
|**[ContactItem](contactitem-object-outlook.md)**|**[ReportItem](reportitem-object-outlook.md)**|
|**[DistListItem](distlistitem-object-outlook.md)**|**[SharingItem](sharingitem-object-outlook.md)**|
|**[DocumentItem](documentitem-object-outlook.md)**|**[TaskItem](taskitem-object-outlook.md)**|
|**[JournalItem](journalitem-object-outlook.md)**|**[TaskRequestAcceptItem](taskrequestacceptitem-object-outlook.md)**|
|**[MailItem](mailitem-object-outlook.md)**|**[TaskRequestDeclineItem](taskrequestdeclineitem-object-outlook.md)**|
|**[MeetingItem](meetingitem-object-outlook.md)**|**[TaskRequestItem](taskrequestitem-object-outlook.md)**|
|**[NoteItem](noteitem-object-outlook.md)**|**[TaskRequestUpdateItem](taskrequestupdateitem-object-outlook.md)**|
|**[PostItem](Outlook.PostItem.md)**||

## Example

The following Visual Basic for Applications (VBA) example displays a dialog box containing the names and identifiers for each  **Category** object contained in the **[Categories](Outlook.NameSpace.Categories.md)** collection associated with the default **[NameSpace](Outlook.NameSpace.md)** object.


```
Private Sub ListCategoryIDs() 
 
 Dim objNameSpace As NameSpace 
 
 Dim objCategory As Category 
 
 Dim strOutput As String 
 
 
 
 ' Obtain a NameSpace object reference. 
 
 Set objNameSpace = Application.GetNamespace("MAPI") 
 
 
 
 ' Check if the Categories collection for the Namespace 
 
 ' contains one or more Category objects. 
 
 If objNameSpace.Categories.Count > 0 Then 
 
 
 
 ' Enumerate the Categories collection. 
 
 For Each objCategory In objNameSpace.Categories 
 
 
 
 ' Add the name and ID of the Category object to 
 
 ' the output string. 
 
 strOutput = strOutput &amp; objCategory.Name &amp; _ 
 
 ": " &amp; objCategory.CategoryID &amp; vbCrLf 
 
 Next 
 
 End If 
 
 
 
 ' Display the output string. 
 
 MsgBox strOutput 
 
 
 
 ' Clean up. 
 
 Set objCategory = Nothing 
 
 Set objNameSpace = Nothing 
 
 
 
End Sub 
 

```


## Properties



|**Name**|
|:-----|
|[Application](Outlook.Category.Application.md)|
|[CategoryBorderColor](Outlook.Category.CategoryBorderColor.md)|
|[CategoryGradientBottomColor](Outlook.Category.CategoryGradientBottomColor.md)|
|[CategoryGradientTopColor](Outlook.Category.CategoryGradientTopColor.md)|
|[CategoryID](Outlook.Category.CategoryID.md)|
|[Class](category-class-property-outlook.md)|
|[Color](Outlook.Category.Color.md)|
|[Name](Outlook.Category.Name.md)|
|[Parent](category-parent-property-outlook.md)|
|[Session](category-session-property-outlook.md)|
|[ShortcutKey](Outlook.Category.ShortcutKey.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
