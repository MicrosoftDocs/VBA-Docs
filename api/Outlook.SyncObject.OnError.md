---
title: SyncObject.OnError event (Outlook)
keywords: vbaol11.chm113
f1_keywords:
- vbaol11.chm113
ms.prod: outlook
api_name:
- Outlook.SyncObject.OnError
ms.assetid: 1faa9708-959c-735b-b6ba-5a78e5fb2690
ms.date: 06/08/2017
localization_priority: Normal
---


# SyncObject.OnError event (Outlook)

Occurs when Microsoft Outlook encounters an error while synchronizing a user's folders using the specified **Send\Receive** group.


## Syntax

_expression_. `OnError`( `_Code_` , `_Description_` )

_expression_ A variable that represents a [SyncObject](Outlook.SyncObject.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Code_|Required| **Long**|A unique value that identifies the error.|
| _Description_|Required| **String**|A textual description of the error.|

## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

This Microsoft Visual Basic for Applications (VBA) example displays a message box describing the synchronization error when an error occurs during synchronization. The sample code must be placed in a class module, and the  `Initialize_handler` routine must be called before the event procedure can be called by Outlook.


```vb
Public WithEvents mySync As Outlook.SyncObject 
 
Sub Initialize_handler() 
 Set mySync = Application.Session.SyncObjects.Item(1) 
 mySync.Start 
 mySync.Stop 
End Sub 
 
Private Sub mySync_OnError(ByVal Code As Long, ByVal Description As String) 
 MsgBox "Unexpected sync error" & Code & ": " & Description 
End Sub
```


## See also


[SyncObject Object](Outlook.SyncObject.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]