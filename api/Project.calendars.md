---
title: Calendars object (Project)
ms.prod: project-server
ms.assetid: a96c7b96-f0ab-5ec3-3d16-facea61b8ee5
ms.date: 06/08/2017
localization_priority: Normal
---


# Calendars object (Project)

Contains a collection of **[Calendar](Project.Calendar.md)** objects.


## Example

 **Using the Calendar Object**

Use **BaseCalendars** (_index_), where _index_ is the calendar index number or calendar name, to return a single **Calendar** object.




```vb
MsgBox ActiveProject.BaseCalendars(1).Name
```

 **Using the Calendars Collection**

Use the **[BaseCalendars](./Project.Project.BaseCalendars.md)** property to return a **Calendars** collection. The following example resets the properties of each base calendar in the active project to their default values.




```vb
Dim C As Calendar 

 

For Each C In ActiveProject.BaseCalendars 

 C.Reset 

Next C
```

Use the **[BaseCalendarCreate](./Project.Application.BaseCalendarCreate.md)** method to add a **Calendar** object to the **Calendars** collection. The following example creates a new base calendar.




```vb
BaseCalendarCreate Name:="Base Holiday Calendar"
```


## Properties



|Name|
|:-----|
|[Application](./Project.Calendars.Application.md)|
|[Count](./Project.Calendars.Count.md)|
|[Item](./Project.Calendars.Item.md)|
|[Parent](./Project.Calendars.Parent.md)|

## See also


[Project Object Model](../project/Concepts/project-object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]