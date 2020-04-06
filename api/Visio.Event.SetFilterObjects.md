---
title: Event.SetFilterObjects method (Visio)
keywords: vis_sdr.chm12650835
f1_keywords:
- vis_sdr.chm12650835
ms.prod: visio
api_name:
- Visio.Event.SetFilterObjects
ms.assetid: 6aa63a44-de34-6cc8-88b2-386064582416
ms.date: 06/08/2017
localization_priority: Normal
---


# Event.SetFilterObjects method (Visio)

Specifies an array of object types and a  **True** or **False** value indicating how to filter events for each object.


## Syntax

_expression_. `SetFilterObjects`( `_Objects()_` )

_expression_ A variable that represents an **[Event](Visio.Event.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Objects()_|Required| **Long**|An array of objects types and a  **True** or **False** value specifying how to filter events for each object type.|

## Return value

Nothing


## Remarks

When an **Event** object created with the **AddAdvise** method is added to the **EventList** collection of a source object, the default behavior is that all occurrences of that event are passed to the event sink. The **SetFilterObjects** method provides a way to ignore selected events based on object type.

The  _Objects()_ parameter passed to **SetFilterObjects** is an array defined in the following manner.

The number of elements in the array is a multiple of 2:




- The first element contains an object type (one of  **visTypePage**, **visTypeGroup**, **visTypeShape**, **visTypeForeignObject**, **visTypeGuide**, or **visTypeDoc**).
    
- The second element contains a  **True** or **False** value indicating whether you are listening to events for that object (**True** to listen to an object's events; **False** to exclude an object's events).
    


For an event to successfully pass through an object event filter, it must satisfy the following criteria:




- It must be a valid object type.
    
- If all filters are  **True**, the event must match at least one filter.
    
- If all filters are  **False**, the event must not match any filter.
    
- If the filters are a mixture of  **True** and **False**, the event must match at least one **True** filter and not match any **False** filters.
    


If there are no  **True** ranges defined in the array, events are considered **True**.

For example, if you want to listen only to events sourced by a shape or guide, you can pass an array like the following:




```vb
 
 Dim aFilterObjects(1 To (2 * 2)) As Long 
 aFilterObjects(1) = visTypeShape 
 aFilterObjects(2) = True 
 aFilterObjects(3) = visTypeGuide 
 aFilterObjects(4) = True 

```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]