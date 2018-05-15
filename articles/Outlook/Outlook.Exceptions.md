---
title: Exceptions Object (Outlook)
keywords: vbaol11.chm289
f1_keywords:
- vbaol11.chm289
ms.prod: outlook
api_name:
- Outlook.Exceptions
ms.assetid: fa3b6c2e-33b0-0f04-4e60-af2c582f2caa
ms.date: 06/08/2017
---


# Exceptions Object (Outlook)

Contains a group of  **[Exception](Outlook.Exception.md)** objects.


## Remarks

If you have a recurring  **[AppointmentItem](Outlook.AppointmentItem.md)**, the **[RecurrencePattern](Outlook.RecurrencePattern.md)** object defines the recurrence of these appointments. The **Exceptions** object contains the group of **Exception** objects that define the exceptions to that series of appointments.

 **Exception** objects are added to the **Exceptions** object whenever a property in the corresponding **AppointmentItem** object is altered.


## Example

The following example sets a reference to the  **Exceptions** object.


```
Set myExceptions = myRecurrencePattern.Exceptions
```


## Methods



|**Name**|
|:-----|
|[Item](Outlook.Exceptions.Item.md)|

## Properties



|**Name**|
|:-----|
|[Application](Outlook.Exceptions.Application.md)|
|[Class](Outlook.Exceptions.Class.md)|
|[Count](Outlook.Exceptions.Count.md)|
|[Parent](Outlook.Exceptions.Parent.md)|
|[Session](Outlook.Exceptions.Session.md)|

## See also


#### Other resources


[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
