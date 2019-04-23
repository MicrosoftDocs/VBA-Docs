---
title: IAssistance.SearchHelp method (Office)
ms.prod: office
api_name:
- Office.IAssistance.SearchHelp
ms.assetid: 807128e9-5125-1650-d53f-cbd50d3e318a
ms.date: 01/16/2019
localization_priority: Normal
---


# IAssistance.SearchHelp method (Office)

Performs a search from the Office Help Viewer based on one or more keywords. Keywords can be a word or a phrase.


## Syntax

_expression_.**SearchHelp** (_Query_, _Scope_)

_expression_ An expression that returns an **[IAssistance](Office.IAssistance.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Query_|Required|**String**|Represents the search keyword or phrase.|
| _Scope_|Optional|**String**|The namespace registered within the host application.|

## Remarks

The **Assistance** property returns an **IAssistance** object. The **IAssistance** object exposes methods that allow developers to display help topics in the Office Help Viewer or to display help topics that ship with Office in the Help window of the host application. Developers either pass specific Help IDs to the help system or pass specific search queries. Help IDs have to be explicitly added to the Help file in order for the Help ID to return the help topic.

The user can return more relevant help by narrowing the scope of the search, as long as the specified scope is applicable to the application. 

The following scopes are available within Microsoft Office applications. By default, the scope is set to the current application's namespace if a **Null** string ("") is passed as a parameter.


## Example

In the first example, the search for "print a document" will be in the "Excel" namespace. 

In the second example, the search for "Application" will be in the "Excel Developer" namespace.


```vb
Sub SearchForHelpTopics1() 
 Application.Assistance.SearchHelp "print a document", "" 
End Sub 
 
Sub SearchForHelpTopics2() 
 Application.Assistance.SearchHelp "Application", "DEV" 
End Sub
```


## See also

- [IAssistance object members](overview/Library-Reference/iassistance-members-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]