---
title: MailItem.Display method (Outlook)
keywords: vbaol11.chm1323
f1_keywords:
- vbaol11.chm1323
ms.prod: outlook
api_name:
- Outlook.MailItem.Display
ms.assetid: 19ead642-b7bd-579f-e43b-ef5c5d0cfecb
ms.date: 06/08/2017
localization_priority: Normal
---


# MailItem.Display method (Outlook)

Displays a new **[Inspector](Outlook.Inspector.md)** object for the item.


## Syntax

_expression_. `Display`( `_Modal_` )

_expression_ A variable that represents a [MailItem](Outlook.MailItem.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Modal_|Optional| **Variant**| **True** to make the window modal. The default value is **False**.|

## Remarks

The **Display** method is supported for explorer and inspector windows for the sake of backward compatibility. To activate an explorer or inspector window, use the **[Activate](Outlook.Inspector.Activate(method).md)** method.

If you attempt to open an "unsafe" file system object (or "freedoc" file) by using the Microsoft Outlook object model, you receive the  **E_FAIL** return code in the C or C++ programming languages. In Outlook 2000 and earlier, you could open an "unsafe" file system object by using the **Display** method.


## Example

This Visual Basic for Applications example displays the first item in the  **Inbox** folder. This example will return an error if the **Inbox** is empty, because you are trying to display a specific item. If there are no items in the folder, a message box will be displayed to inform the user.


> [!NOTE] 
> The items in the  **Items** collection object are not guaranteed to be in any particular order.


```vb
Sub DisplayFirstItem() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 
 Set myFolder = myNameSpace.GetDefaultFolder(olFolderInbox) 
 
 On Error GoTo ErrorHandler 
 
 myFolder.Items(1).Display 
 
 Exit Sub 
 
ErrorHandler: 
 
 MsgBox "There are no items to display." 
 
End Sub
```


## See also


[MailItem Object](Outlook.MailItem.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
