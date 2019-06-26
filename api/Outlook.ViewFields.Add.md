---
title: ViewFields.Add method (Outlook)
keywords: vbaol11.chm2552
f1_keywords:
- vbaol11.chm2552
ms.prod: outlook
api_name:
- Outlook.ViewFields.Add
ms.assetid: 0bf96999-fdb8-d13c-6409-cee150a32c06
ms.date: 06/08/2017
localization_priority: Normal
---


# ViewFields.Add method (Outlook)

Adds the specified field to the end of the **[ViewFields](Outlook.ViewFields.md)** collection for the view.


## Syntax

_expression_.**Add** (_PropertyName_)

_expression_ A variable that represents a [ViewFields](Outlook.ViewFields.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _PropertyName_|Required| **String**|The name of the property to which the new object is associated. This property can be referenced by field name (displayed in the Field Chooser) or by namespace (represented by **[ViewField.ViewXMLSchemaName](Outlook.ViewField.ViewXMLSchemaName.md)**).|

## Return value

A **ViewField** object that represents the new view field.


## Remarks

To programmatically add a custom field to a view, use the **ViewFields.Add** method. This is the recommended way to dynamically change the view over setting the **[XML](Outlook.View.XML.md)** property of the **[View](Outlook.View.md)** object.

Referencing the property in  _PropertyName_ by its field name requires the localized name in the corresponding locale. For more information on referencing properties by namespace, see [Referencing Properties by Namespace](../outlook/How-to/Navigation/referencing-properties-by-namespace.md).

If you are adding a custom property to the **ViewFields** collection, the property must exist in the **[UserDefinedProperties](Outlook.Folder.UserDefinedProperties.md)** collection for the View's parent folder.

If the property already exists in the **ViewFields** collection, Outlook will raise an error.

Certain properties cannot be added to a view using **ViewFields.Add**, including binary properties, computed properties, and HTML or RTF body content. For more information, see [Unsupported Properties in a Table Object or Table Filter](../outlook/How-to/Search-and-Filter/unsupported-properties-in-a-table-object-or-table-filter.md).


## Example

The following code sample in Microsoft Visual Basic for Applications (VBA) adds the Subject field to the current view of the Inbox, referencing it by its field name. To avoid Outlook raising an error, it tests for the presence of the field in the **ViewFields** collection representing the current view of the Inbox before adding it.


```vb
Sub DemoViewFieldsAdd() 
 
 Dim oTableView As Outlook.TableView 
 
 Dim oViewFields As Outlook.ViewFields 
 
 Dim oViewField As Outlook.ViewField 
 
 Dim oInbox As Outlook.folder 
 
 
 
 Set oInbox = Application.Session.GetDefaultFolder(olFolderInbox) 
 
 
 
 On Error GoTo Err_Handler 
 
 
 
 If oInbox.CurrentView.ViewType = olTableView Then 
 
 Set oTableView = oInbox.CurrentView 
 
 Set oViewField = oTableView.ViewFields("Subject") 
 
 If oViewField Is Nothing Then 
 
 Set oViewField = oTableView.ViewFields.Add("Subject") 
 
 End If 
 
 End If 
 
 Exit Sub 
 
 
 
Err_Handler: 
 
 MsgBox Err.Description, vbExclamation 
 
 Resume Next 
 
End Sub
```

The following code sample in VBA assumes the current view is a **[TableView](Outlook.TableView.md)**, references the Message Class property by namespace and adds it to the current view of the current folder. To avoid Outlook raising an error, the code checks for the existence of this property in the view before calling **ViewFields.Add**.




```vb
Sub ViewFieldsAdd() 
 
 Dim oFolder As Outlook.Folder 
 
 Dim oView As Outlook.TableView 
 
 Dim oViewField As Outlook.ViewField 
 
 On Error Resume Next 
 
 Dim PR_MESSAGE_CLASS As String 
 
 PR_MESSAGE_CLASS = "http://schemas.microsoft.com/mapi/proptag/0x001a001e" 
 
 Set oFolder = Application.ActiveExplorer.CurrentFolder 
 
 If oFolder.CurrentView.ViewType = olTableView Then 
 
 Set oView = oFolder.CurrentView 
 
 'Determine if the ViewField exists in ViewFields collection 
 
 If oView.ViewFields(PR_MESSAGE_CLASS) Is Nothing Then 
 
 Set oViewField = oView.ViewFields.Add(PR_MESSAGE_CLASS) 
 
 'Persist the changes 
 
 oView.Save 
 
 End If 
 
 End If 
 
End Sub
```


## See also


[ViewFields Object](Outlook.ViewFields.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]