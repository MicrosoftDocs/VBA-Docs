---
title: IAssistance.ClearDefaultContext method (Office)
keywords: vbaof11.chm326004
f1_keywords:
- vbaof11.chm326004
ms.prod: office
api_name:
- Office.IAssistance.ClearDefaultContext
ms.assetid: ebdc0b7e-f459-6d4d-af45-0e5625b2448e
ms.date: 01/16/2019
localization_priority: Normal
---


# IAssistance.ClearDefaultContext method (Office)

Clears the default help topic previously defined in the **[SetDefaultContext](office.iassistance.setdefaultcontext.md)** method.


## Syntax

_expression_.**ClearDefaultContext** (_HelpId_)

_expression_ An expression that returns an **[IAssistance](Office.IAssistance.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _HelpId_|Required|**String**|The ID of the default help topic.|

## Remarks

Executing this method will stop the default help topic from displaying when the user presses **F1** or chooses the **Help** button in a dialog box.

The **Assistance** property returns an **IAssistance** object. The **IAssistance** object exposes methods that allow developers to display help topics in the Office Help Viewer or to display help topics that ship with Office in the Help window of the host application. Developers either pass specific Help IDs to the help system or pass specific search queries. Help IDs have to be explicitly added to the Help file in order for the Help ID to return the help topic.


## Example

In the following example, the default help topic is cleared and will no longer be displayed.


```vb
Sub ClearDefaultHelpTopic() 
 Application.Assistance.ClearDefaultContext "22261" 
End Sub
```


## See also

- [IAssistance object members](overview/Library-Reference/iassistance-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]