---
title: Items.FindNext method (Outlook)
keywords: vbaol11.chm63
f1_keywords:
- vbaol11.chm63
ms.prod: outlook
api_name:
- Outlook.Items.FindNext
ms.assetid: 2530f640-e024-3567-f539-6bdbf645401d
ms.date: 06/08/2017
localization_priority: Normal
---


# Items.FindNext method (Outlook)

After the  **[Find](Outlook.Items.Find.md)** method runs, this method finds and returns the next Outlook item in the specified collection.


## Syntax

_expression_.**FindNext**

_expression_ A variable that represents an [Items](Outlook.Items.md) object.


## Return value

An **Object** value that represents the next Outlook item found in the collection.


## Remarks

 The search operation begins from the current position, which matches the expression previously set through the **Find** method.

The method returns an Outlook item object if the call succeeds; it returns  **Null** (or **Nothing** in Visual Basic) if it fails.


## Example

This Visual Basic for Applications (VBA) example uses the  **[GetDefaultFolder](Outlook.NameSpace.GetDefaultFolder.md)** method to return the **[Folder](Outlook.Folder.md)** object that represents the default **Calendar** folder for the current user. It then uses the **[Find](Outlook.Items.Find.md)** and **FindNext** methods to locate all the appointments that occur today and display them in a series of message boxes.


```vb
Sub DemoFindNext() 
 Dim myNameSpace As Outlook.NameSpace 
 Dim tdystart As Date 
 Dim tdyend As Date 
 Dim myAppointments As Outlook.Items 
 Dim currentAppointment As Outlook.AppointmentItem 
 
 Set myNameSpace = Application.GetNamespace("MAPI") 
 tdystart = VBA.Format(Now, "Short Date") 
 tdyend = VBA.Format(Now + 1, "Short Date") 
 Set myAppointments = myNameSpace.GetDefaultFolder(olFolderCalendar).Items 
 Set currentAppointment = myAppointments.Find("[Start] >= """ & tdystart & """ and [Start] <= """ & tdyend & """") 
 While TypeName(currentAppointment) <> "Nothing" 
 MsgBox currentAppointment.Subject 
 Set currentAppointment = myAppointments.FindNext 
Wend 
End Sub
```


## See also


[Items Object](Outlook.Items.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]