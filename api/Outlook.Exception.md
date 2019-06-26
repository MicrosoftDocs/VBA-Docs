---
title: Exception object (Outlook)
keywords: vbaol11.chm296
f1_keywords:
- vbaol11.chm296
ms.prod: outlook
api_name:
- Outlook.Exception
ms.assetid: 010552b0-9ba6-c81b-1e3a-fd6a681e5163
ms.date: 06/08/2017
localization_priority: Normal
---


# Exception object (Outlook)

Represents information about one instance of an  **[AppointmentItem](Outlook.AppointmentItem.md)** object which is an exception to a recurring series.


## Remarks

 Unlike most of the other Microsoft Outlook objects, the **Exception** object is a read-only object. This means that you cannot create an **Exception** object but, rather, the object is created when a property of an **AppointmentItem** is altered. For example, if you change the **[Start](Outlook.AppointmentItem.Start.md)** property of one **AppointmentItem**, you have created an **Exception** in **AppointmentItem.RecurrencePattern.Exceptions**.


> [!NOTE] 
> The  **[Exceptions](Outlook.Exceptions.md)** object is on the **[RecurrencePattern](Outlook.RecurrencePattern.md)**, not the **AppointmentItem** object itself.

When you work with recurring appointment items, you should release any prior references, obtain new references to the recurring appointment item before you access or modify the item, and release these references as soon as you are finished and have saved the changes. This practice applies to the recurring  **AppointmentItem** object, and any **[Exception](Outlook.Exception.md)** or **RecurrencePattern** object. To release a reference in Visual Basic for Applications (VBA) or Visual Basic, set that existing object to **Nothing**. In C#, explicitly release the memory for that object. For a code example, see the topic for the **AppointmentItem** object.

Note that even after you release your reference and attempt to obtain a new reference, if there is still an active reference, held by another add-in or Outlook, to one of the above objects, your new reference will still point to an out-of-date copy of the object. Therefore, it is important that you release your references as soon as you are finished with the recurring appointment.


## Example

The following Visual Basic for Applications (VBA) example retrieves the first  **Exception** object from the **Exceptions** collection object associated with a **RecurrencePattern** object.


```vb
Sub GetException() 
 
 Dim myNameSpace As Outlook.NameSpace 
 
 Dim myFolder As Outlook.Folder 
 
 Dim myItems As Outlook.Items 
 
 Dim myApptItem As Outlook.AppointmentItem 
 
 Dim myRecurrencePattern As Outlook.RecurrencePattern 
 
 Dim myException As Outlook.Exception 
 
 
 
 Set myNameSpace = Application.GetNameSpace("MAPI") 
 
 Set myFolder = myNameSpace.GetDefaultFolder(olFolderCalendar) 
 
 Set myItems = myFolder.Items 
 
 Set myApptItem = myItems("Daily Meeting") 
 
 Set myRecurrencePattern = myApptItem.GetRecurrencePattern 
 
 Set myException = myRecurrencePattern.Exceptions.Item(1) 
 
End Sub
```


## Properties



|Name|
|:-----|
|[Application](Outlook.Exception.Application.md)|
|[AppointmentItem](Outlook.Exception.AppointmentItem.md)|
|[Class](Outlook.Exception.Class.md)|
|[Deleted](Outlook.Exception.Deleted.md)|
|[OriginalDate](Outlook.Exception.OriginalDate.md)|
|[Parent](Outlook.Exception.Parent.md)|
|[Session](Outlook.Exception.Session.md)|

## See also


[Outlook Object Model Reference](overview/Outlook/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]