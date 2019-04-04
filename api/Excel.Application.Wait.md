---
title: Application.Wait method (Excel)
keywords: vbaxl10.chm133242
f1_keywords:
- vbaxl10.chm133242
ms.prod: excel
api_name:
- Excel.Application.Wait
ms.assetid: 71425d1c-6b37-a510-d8b5-072136e98f04
ms.date: 04/05/2019
localization_priority: Priority
---


# Application.Wait method (Excel)

Pauses a running macro until a specified time. Returns **True** if the specified time has arrived.


## Syntax

_expression_.**Wait** (_Time_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Time_|Required| **Variant**|The time at which you want the macro to resume, in Microsoft Excel date format.|


## Return value

**Boolean**


## Remarks

The **Wait** method suspends all Microsoft Excel activity and may prevent you from performing other operations on your computer while **Wait** is in effect. However, background processes such as printing and recalculation continue.


## Example

This example pauses a running macro until 6:23 P.M. today.

```vb
Application.Wait "18:23:00"
```

<br/>

This example pauses a running macro for approximately 10 seconds.

```vb
newHour = Hour(Now()) 
newMinute = Minute(Now()) 
newSecond = Second(Now()) + 10 
waitTime = TimeSerial(newHour, newMinute, newSecond) 
Application.Wait waitTime
```

<br/>

This example displays a message indicating whether 10 seconds have passed.

```vb
If Application.Wait(Now + TimeValue("0:00:10")) Then 
 MsgBox "Time expired" 
End If
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
