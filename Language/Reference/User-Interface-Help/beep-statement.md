---
title: Beep statement (VBA)
keywords: vblr6.chm1008861
f1_keywords:
- vblr6.chm1008861
ms.prod: office
ms.assetid: 61328fce-c26c-2758-436a-474da9aac8b7
ms.date: 12/03/2018
localization_priority: Priority
---


# Beep statement

Sounds a tone through the computer's speaker.

## Syntax

**Beep**

## Remarks

The frequency and duration of the beep depend on your hardware and system software, and vary among computers.

## Example

This example uses the **Beep** statement to sound three consecutive tones through the computer's speaker.


```vb
Dim I 
For I = 1 To 3 ' Loop 3 times. 
 Beep ' Sound a tone. 
Next I 

```

## See also

- [Data types](data-type-summary.md)
- [Statements](../statements.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]