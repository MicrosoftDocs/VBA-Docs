---
title: ViewCtl Members (Outlook View Control)
ms.prod: outlook
ms.assetid: 32df30fd-d02c-30c4-7474-0dc359f99f46
ms.date: 06/08/2017
localization_priority: Normal
---


# ViewCtl Members (Outlook View Control)

The Microsoft Outlook View Control displays information about a specific folder and can be integrated into solutions that provide access to Outlook data. The  **ViewCtl** object provides programmatic access to the View Control. The control can be placed in any container that supports ActiveX® controls, including an HTML page that is hosted in Outlook as a Folder Home Page, or a custom Outlook form. If the View Control is placed in an HTML page that is hosted in a browser such as Internet Explorer, some functions of the control are disabled for security.


## Methods



|**Name**|**Description**|
|:-----|:-----|
| **[AddressBook](../../../api/Outlook.viewctl.addressbo.md)**|Displays the Microsoft Outlook  **Address Book** dialog box.|
| **[AddToPFFavorites](../../../api/Outlook.viewctl.addtopffavorit.md)**|Adds the current public folder to the user's Microsoft Exchange Server  **Favorites** public folder.|
| **[AdvancedFind](../../../api/Outlook.viewctl.advancedfi.md)**|Displays the Microsoft Outlook  **Advanced Find** dialog box.|
| **[Categories](../../../api/Outlook.viewctl.categori.md)**|Displays the Microsoft Outlook  **Categories** dialog box for the currently selected item or items in the control.|
| **[CollapseAllGroups](../../../api/Outlook.viewctl.collapseallgrou.md)**|Collapses (closes) all groups that are displayed in the control.|
| **[CollapseGroup](../../../api/Outlook.viewctl.collapsegro.md)**|Collapses (closes) the group that is currently selected in the control. |
| **[CustomizeView](../../../api/Outlook.viewctl.customizevi.md)**|Displays the Microsoft Outlook  **View Summary** dialog box.|
| **[Delete](../../../api/Outlook.viewctl.dele.md)**|After prompting the user to confirm, deletes the groups or items that are currently selected in the control. |
| **[ExpandAllGroups](../../../api/Outlook.viewctl.expandallgrou.md)**|Expands (opens) all groups that are displayed in the control. |
| **[ExpandGroup](../../../api/Outlook.viewctl.expandgro.md)**|Expands (opens) the group that is currently selected in the control. |
| **[FlagItem](../../../api/Outlook.viewctl.flagit.md)**|Displays the Microsoft Outlook  **Flag for Follow Up** dialog box for the selected item.|
| **[ForceUpdate](../../../api/Outlook.viewctl.forceupda.md)**|Refreshes the view in the control, applying any property changes made since the  **[DeferUpdate](../../../api/Outlook.viewctl.deferupda.md)** property was set to **True**.|
| **[Forward](../../../api/Outlook.viewctl.forwa.md)**|Executes the Forward action for the item or items that are selected in the control.|
| **[GoToDate](../../../api/Outlook.viewctl.gotoda.md)**|Opens a calendar view of a specific date.|
| **[NewAppointment](../../../api/Outlook.viewctl.newappointme.md)**|Creates and displays a new appointment.|
| **[NewContact](../../../api/Outlook.viewctl.newconta.md)**|Creates and displays a new contact.|
| **[NewDefaultItem](../../../api/Outlook.viewctl.newdefaultit.md)**|Creates and displays a new Microsoft Outlook item. |
| **[NewDistributionList](../../../api/Outlook.viewctl.newdistributionli.md)**|Creates and displays a new distribution list.|
| **[NewForm](../../../api/Outlook.viewctl.newfo.md)**|Displays the Microsoft Outlook  **Choose Form** dialog box.|
| **[NewJournalEntry](../../../api/Outlook.viewctl.newjournalent.md)**|Creates and displays a new journal entry.|
| **[NewMeetingRequest](../../../api/Outlook.viewctl.newmeetingreque.md)**|Creates and displays a new meeting request.|
| **[NewMessage](../../../api/Outlook.viewctl.newmessa.md)**|Creates and displays a new email message.|
| **[NewNote](../../../api/Outlook.viewctl.newno.md)**|Creates and displays a new note item.|
| **[NewPost](../../../api/Outlook.viewctl.newpo.md)**|Creates and displays a new post item.|
| **[NewTask](../../../api/Outlook.viewctl.newta.md)**|Creates and displays a new task.|
| **[NewTaskRequest](../../../api/Outlook.viewctl.newtaskreque.md)**|Creates and displays a new task request.|
| **[Open](../../../api/Outlook.viewctl.op.md)**|Opens the item or items that are currently selected in the control.|
| **[OpenSharedDefaultFolder](../../../api/Outlook.viewctl.openshareddefaultfold.md)**|Displays a specified user's default folder in the control.|
| **[PrintItem](../../../api/Outlook.viewctl.printit.md)**|Prints the items that are currently selected in the control. |
| **[Reply](../../../api/Outlook.viewctl.rep.md)**|Executes the Reply action for the item or items selected in the control.|
| **[ReplyAll](../../../api/Outlook.viewctl.replya.md)**|Executes the ReplyAll action for the item or items that are selected in the control.|
| **[ReplyInFolder](../../../api/Outlook.viewctl.replyinfold.md)**|Creates a post item for each message that is currently selected in the control.|
| **[SaveAs](../../../api/Outlook.viewctl.save.md)**|Saves the items that are selected in the control as a single file.|
| **[SendAndReceive](../../../api/Outlook.viewctl.sendandrecei.md)**|Sends all messages that are in the  **Outbox** folder and checks for new messages.|
| **[ShowFields](../../../api/Outlook.viewctl.showfiel.md)**|Displays the Microsoft Outlook  **Show Fields** dialog box.|
| **[Sort](../../../api/Outlook.viewctl.so.md)**|Displays the Microsoft Outlook  **Sort** dialog box.|
| **[SynchFolder](../../../api/Outlook.viewctl.synchfold.md)**|Synchronizes the online and offline folders that are displayed in the control. |

## Properties



|**Name**|**Description**|
|:-----|:-----|
| **[ActiveFolder](../../../api/Outlook.viewctl.activefold.md)**|Returns an object that represents the folder displayed in the control. Read-only.|
| **[DeferUpdate](../../../api/Outlook.viewctl.deferupda.md)**|Gets or sets a  **Boolean** value that indicates whether property changes affect the control display. Read/write.|
| **[EnableRowPersistance](../../../api/Outlook.viewctl.enablerowpersistan.md)**|Gets or sets a value that indicates whether the View Control retains state information about the last selected row. Read/write.|
| **[Filter](../../../api/Outlook.viewctl.filt.md)**|Gets or sets a  **String** that represents the Distributed Authoring and Versioning (DAV) Searching and Locating (DASL) statement used to restrict the display to a specified subset of data. Read/write.|
| **[FilterAppend](../../../api/Outlook.viewctl.filterappe.md)**|Gets or sets a  **String** that represents the additional criteria to add to the filter settings. Read/write.|
| **[Folder](../../../api/Outlook.viewctl.fold.md)**|Gets or sets a  **String** that represents the path of the folder displayed by the control.|
| **[ItemCount](../../../api/Outlook.viewctl.itemcou.md)**|Returns a  **Long** that indicates the count of objects in the current folder displayed in the control. Read-only.|
| **[Namespace](../../../api/Outlook.viewctl.namespa.md)**|Returns or sets a  **String** that represents the namespace property of the control. Read/write.|
| **[OutlookApplication](../../../api/Outlook.viewctl.outlookapplicati.md)**|Returns an object that represents the container object for the control. Read-only.|
| **[Restriction](../../../api/Outlook.viewctl.restricti.md)**|Sets or returns a  **String** that represents a filter to the items that are displayed in the control. As a result, the control displays only those items that match the filter. Read/write.|
| **[SelectedDate](../../../api/Outlook.viewctl.selectedda.md)**|Returns or sets the selected date. Read-only.|
| **[Selection](../../../api/Outlook.viewctl.selecti.md)**|Returns a  **[Selection](../../../api/Outlook.Selection.md)** object that consists of one or more items that are selected in the current view. Read-only.|
| **[View](../../../api/Outlook.viewctl.vi.md)**|Returns or sets a  **String** that represents the name of the view in the control. Read/write.|
| **[ViewXML](../../../api/Outlook.viewctl.viewx.md)**|Returns or sets a  **String** that represents the view implementation via XML. Read/write.|

## Events



|**Name**|**Description**|
|:-----|:-----|
| **[Activate](../../../api/Outlook.viewctl.activa.md)**|Occurs when a View Control becomes the active element on the page, either as a result of user action or through program code.|
| **[BeforeViewSwitch](../../../api/Outlook.viewctl.beforeviewswit.md)**|Occurs before Microsoft Outlook changes the view that is applied to the folder displayed in the View Control element, either as a result of user action or through program code. |
| **[SelectionChange](../../../api/Outlook.viewctl.selectionchan.md)**|Occurs when the selection of the current view changes. |
| **[ViewSwitch](../../../api/Outlook.viewctl.viewswit.md)**|Occurs when Microsoft Outlook changes the view that is applied to the folder displayed in the View Control element, either as a result of user action or through program code.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]