---
title: Application.GetNamespace method (Outlook)
keywords: vbaol11.chm717
f1_keywords:
- vbaol11.chm717
ms.prod: outlook
api_name:
- Outlook.Application.GetNamespace
ms.assetid: 6175d0d9-5a61-ce45-35c0-b70895d757b3
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.GetNamespace method (Outlook)

Returns a **[NameSpace](Outlook.NameSpace.md)** object of the specified type.


## Syntax

_expression_. `GetNamespace`( `_Type_` )

_expression_ A variable that represents an **[Application](Outlook.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **String**|The type of name space to return.|

## Return value

A **NameSpace** object that represents the specified namespace.


## Remarks

The only supported name space type is "MAPI". The **GetNameSpace** method is functionally equivalent to the **Session** property.


## Example

This Visual Basic for Applications (VBA) example uses the  **[CurrentFolder](Outlook.Explorer.CurrentFolder.md)** property to change the displayed folder to the user's **Calendar** folder.


```vb
Sub ChangeCurrentFolder() 
 
 Dim myNamespace As Outlook.NameSpace 
 
 
 
 Set myNamespace = Application.GetNamespace("MAPI") 
 
 Set Application.ActiveExplorer.CurrentFolder = _ 
 
 myNamespace.GetDefaultFolder(olFolderCalendar) 
 
End Sub
```


## See also


[Application Object](Outlook.Application.md)



[How to: Obtain and Log On to an Instance of Outlook](../outlook/How-to/Security/obtain-and-log-on-to-an-instance-of-outlook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
