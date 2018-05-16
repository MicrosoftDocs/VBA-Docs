---
title: NavigationButton.EventProcPrefix Property (Access)
keywords: vbaac10.chm10447
f1_keywords:
- vbaac10.chm10447
ms.prod: access
api_name:
- Access.NavigationButton.EventProcPrefix
ms.assetid: 84bf1794-9b36-91eb-23d3-e5db4e951f85
ms.date: 06/08/2017
---


# NavigationButton.EventProcPrefix Property (Access)

Gets or sets the prefix portion of an event procedure name. Read/write  **String**.


## Syntax

 _expression_. **EventProcPrefix**

 _expression_ A variable that represents a **NavigationButton** object.


## Remarks

For example, if you have a command button with an event procedure named Details_Click, the  **EventProcPrefix** property returns the string "Details".

Microsoft Access adds the prefix portion of an event procedure name to the event name with an underscore character (_).


## See also


#### Concepts


[NavigationButton Object](Access.NavigationButton.md)

