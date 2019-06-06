---
title: Application.CommandBars property (Publisher)
keywords: vbapb10.chm131088
f1_keywords:
- vbapb10.chm131088
ms.prod: publisher
api_name:
- Publisher.Application.CommandBars
ms.assetid: 21537c04-d406-6016-4f35-2f6ce6851db2
ms.date: 06/04/2019
localization_priority: Normal
---


# Application.CommandBars property (Publisher)

Sets or returns a **[CommandBars](office.commandbars.md)** collection that represents the menu bar and all the toolbars in Microsoft Publisher.


## Syntax

_expression_.**CommandBars**

_expression_ A variable that represents an **[Application](Publisher.Application.md)** object.


## Return value

CommandBars


## Example

This example enlarges all command bar buttons, enables ToolTips, and shows all menu items when displaying menus.

```vb
Sub CmdBars() 
 
 With CommandBars 
 .LargeButtons = False 
 .DisplayTooltips = True 
 .AdaptiveMenus = False 
 End With 
 
End Sub
```

<br/>

This example displays the **Objects** toolbar at the bottom of the application window.

```vb
Sub ShowObjectsToolbar 
 
 With CommandBars("Objects") 
 .Visible = True 
 .Position = msoBarBottom 
 End With 
 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]