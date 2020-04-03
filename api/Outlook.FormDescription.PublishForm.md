---
title: FormDescription.PublishForm method (Outlook)
keywords: vbaol11.chm201
f1_keywords:
- vbaol11.chm201
ms.prod: outlook
api_name:
- Outlook.FormDescription.PublishForm
ms.assetid: 2040736a-4be0-90c4-0dfc-20c6ee4eb305
ms.date: 06/08/2017
localization_priority: Normal
---


# FormDescription.PublishForm method (Outlook)

Saves the definition of the  **[FormDescription](Outlook.FormDescription.md)** object in the specified form registry (library).


## Syntax

_expression_. `PublishForm`( `_Registry_` , `_Folder_` )

_expression_ A variable that represents a [FormDescription](Outlook.FormDescription.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Registry_|Required| **[OlFormRegistry](Outlook.OlFormRegistry.md)**|The form class.|
| _Folder_|Optional| **Variant**|Expression that returns a  **[Folder](Outlook.Folder.md)** object. Used only with Folder form registry. The folder object from which the forms must be accessed.|

## Remarks


> [!NOTE] 
> The  **[Name](Outlook.FormDescription.Name.md)** property must be set before you can use the **PublishForm** method.

Forms are registered as one of three classes: Folder, Organization, or Personal. The Folder form registry holds a set of forms that are only accessible from that specific folder, whether public or private. The Organization form registry holds forms that are shared across an entire enterprise and are accessible to everyone. The Personal form registry holds forms that are accessible only to the current store user.


## Example

This Visual Basic for Applications (VBA) example creates a contact, obtains its  **[FormDescription](Outlook.FormDescription.md)** object, and saves it in the Folder form registry of the default **Contacts** folder.


> [!NOTE] 
> The  **[PublishForm](Outlook.FormDescription.PublishForm.md)** method will return an error if the caption (**[Name](Outlook.FormDescription.Name.md)**) for the form is not set first.


```vb
Sub PublishToFolder() 
 
 Dim myNamespace As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myItem As Outlook.ContactItem 
 
 Dim myForm As Outlook.FormDescription 
 
 
 
 Set myNamespace = Application.GetNamespace("MAPI") 
 
 Set myFolder = _ 
 
 myNamespace.GetDefaultFolder(olFolderContacts) 
 
 Set myItem = Application.CreateItem(olContactItem) 
 
 Set myForm = myItem.FormDescription 
 
 myForm.Name = "My Contact" 
 
 myForm.PublishForm olFolderRegistry, myFolder 
 
End Sub
```

This VBA example creates an appointment, obtains its  **FormDescription** object, and saves it in the user's Personal form registry.



To view the form after you have published it, on the  **File** menu, point to **New**, and click  **Choose Form**. In the **Look in** box, click **Personal Forms Library**. To open your new form, double-click **Interview Scheduler**.




```vb
Set myItem = Application.CreateItem(olAppointmentItem) 
 
Set myForm = myItem.FormDescription 
 
myForm.Name = "Interview Scheduler" 
 
myForm.PublishForm olPersonalRegistry
```


## See also


[FormDescription Object](Outlook.FormDescription.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]