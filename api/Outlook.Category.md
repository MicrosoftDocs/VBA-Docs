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
|**[AppointmentItem](../missing-files/Outlook/appointmentitem-object-outlook.md)**|**[RemoteItem](../missing-files/Outlook/remoteitem-object-outlook.md)**|
|**[ContactItem](../missing-files/Outlook/contactitem-object-outlook.md)**|**[ReportItem](../missing-files/Outlook/reportitem-object-outlook.md)**|
|**[DistListItem](../missing-files/Outlook/distlistitem-object-outlook.md)**|**[SharingItem](../missing-files/Outlook/sharingitem-object-outlook.md)**|
|**[DocumentItem](../missing-files/Outlook/documentitem-object-outlook.md)**|**[TaskItem](../missing-files/Outlook/taskitem-object-outlook.md)**|
|**[JournalItem](../missing-files/Outlook/journalitem-object-outlook.md)**|**[TaskRequestAcceptItem](../missing-files/Outlook/taskrequestacceptitem-object-outlook.md)**|
|**[MailItem](../missing-files/Outlook/mailitem-object-outlook.md)**|**[TaskRequestDeclineItem](../missing-files/Outlook/taskrequestdeclineitem-object-outlook.md)**|
|**[MeetingItem](../missing-files/Outlook/meetingitem-object-outlook.md)**|**[TaskRequestItem](../missing-files/Outlook/taskrequestitem-object-outlook.md)**|
|**[NoteItem](../missing-files/Outlook/noteitem-object-outlook.md)**|**[TaskRequestUpdateItem](../missing-files/Outlook/taskrequestupdateitem-object-outlook.md)**|
|**[PostItem](Outlook.PostItem.md)**||

## Example

The following Visual Basic for Applications (VBA) example displays a dialog box containing the names and identifiers for each  **Category** object contained in the **[Categories](Outlook.NameSpace.Categories.md)** collection associated with the default **[NameSpace](Outlook.NameSpace.md)** object.


```vb
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
|[Class](../missing-files/Outlook/category-class-property-outlook.md)|
|[Color](Outlook.Category.Color.md)|
|[Name](Outlook.Category.Name.md)|
|[Parent](../missing-files/Outlook/category-parent-property-outlook.md)|
|[Session](../missing-files/Outlook/category-session-property-outlook.md)|
|[ShortcutKey](Outlook.Category.ShortcutKey.md)|

## See also


[Outlook Object Model Reference](./overview/object-model-outlook-vba-reference.md)
