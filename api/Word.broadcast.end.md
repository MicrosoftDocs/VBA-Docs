---
title: Broadcast.End method (Word)
keywords: vbawd10.chm36438120
f1_keywords:
- vbawd10.chm36438120
ms.prod: word
ms.assetid: dca52c1c-c337-f9ee-6c82-ef05da5cdf45
ms.date: 06/08/2017
localization_priority: Normal
---


# Broadcast.End method (Word)

Ends the specified broadcast session.


## Syntax

_expression_.**End**

_expression_ A variable that represents a **[Broadcast](Word.broadcast.md)** object.


## Return value

 **VOID**


## Remarks

Calling the **End** method terminates the broadcast session without displaying a confirmation prompt to the user. It also sets the value of the [Broadcast.AttendeeURL](Word.broadcast.attendeeurl.md) property to an empty string.

If the document is not being broadcast, the method returns run-time error 4702.


## See also


[Broadcast Object](Word.broadcast.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]