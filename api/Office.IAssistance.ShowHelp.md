---
title: IAssistance.ShowHelp method (Office)
ms.prod: office
api_name:
- Office.IAssistance.ShowHelp
ms.assetid: 18b46084-114b-69a7-f108-07e4a455e024
ms.date: 01/16/2019
localization_priority: Normal
---


# IAssistance.ShowHelp method (Office)

Displays the help topic specified by its ID in the Office Help Viewer, or for help topics that ship with Office, in the Help window of the current Office application.


## Syntax

_expression_.**ShowHelp** (_HelpId_, _Scope_)

_expression_ An expression that returns an **[IAssistance](Office.IAssistance.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _HelpId_|Optional|**String**|The ID of the help topic.|
| _Scope_|Optional|**String**|The namespace registered within the host application.|

## Remarks

The **Assistance** property returns an **IAssistance** object. The **IAssistance** object exposes methods that allow developers to display help topics in the Office Help Viewer or to display help topics that ship with Office in the Help window of the host application. Developers either pass specific Help IDs to the help system or pass specific search queries. Help IDs have to be explicitly added to the Help file in order for the Help ID to return the help topic.

The following scopes are available within Microsoft Office applications. By default, the scope is set to the current application's namespace if a **Null** string ("") is passed as a parameter.


## Example

In the first line in the following example, the client help viewer displays the help topic associated with the ID `"xlmain11.chm60407"` in the "Excel" namespace. 

In the second line, the client viewer remains open and displays the help topic associated with the ID `"65879"` in the "Excel Developer" namespace. 


```vb
Sub DisplayHelpTopic() 
 Application.Assistance.ShowHelp "xlmain11.chm60407", "" 
 Application.Assistance.ShowHelp "vbaxl10.chm65879", "DEV" 
End Sub
```


## See also

- [IAssistance object members](overview/Library-Reference/iassistance-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]