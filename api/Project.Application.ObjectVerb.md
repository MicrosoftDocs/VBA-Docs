---
title: Application.ObjectVerb method (Project)
keywords: vbapj.chm237
f1_keywords:
- vbapj.chm237
ms.prod: project-server
api_name:
- Project.Application.ObjectVerb
ms.assetid: 55507406-5a36-0361-3b91-7f17860dc577
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ObjectVerb method (Project)

Instructs the active object to perform an action.


## Syntax

_expression_. `ObjectVerb`( `_Verb_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Verb_|Optional|**Long**|The action that the active object should perform.|

## Return value

 **Boolean**


## Remarks

For a list of the actions an object can perform, select the object, and then run the **Object** command.

To determine the number associated with a particular action, run regedit.exe by clicking the Windows  **Start** button and then clicking **Run**. The RegEdit.exe file is in the `%windir%` folder.

Negotiate the registry tree to HKEY_CLASSES_ROOT\  _AppName_. _DocumentName_ \protocol\StdFileEditing\Verb\ _number_, where _AppName_ is the name of the application, _DocumentName_ is the name of the document, and _number_ is the key for an action. For Microsoft Office PowerPoint 2007 , for example, HKEY_CLASSES_ROOT\PowerPoint.Show.12\protocol\StdFileEditing\Verb\0 is the key for the **Show** command.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]