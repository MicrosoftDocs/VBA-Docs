---
title: Application.Macro method (Project)
keywords: vbapj.chm1001
f1_keywords:
- vbapj.chm1001
ms.prod: project-server
api_name:
- Project.Application.Macro
ms.assetid: e07686b6-3c38-7413-692b-aac8fb9bf526
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.Macro method (Project)

Runs a macro.


## Syntax

_expression_. `Macro`( `_Name_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**String**|The name of the macro to run. If  **Name** is omitted, the **Macros** dialog box appears.|

## Return value

 **Boolean**


## Example

The following example runs a macro named "CheckShifts".


```vb
Sub RunMacro() 
 Macro "CheckShifts" 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]