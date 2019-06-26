---
title: Application.AppExecute method (Project)
keywords: vbapj.chm8
f1_keywords:
- vbapj.chm8
ms.prod: project-server
api_name:
- Project.Application.AppExecute
ms.assetid: af263a18-9b88-e6c2-d44c-a2ac41951624
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.AppExecute method (Project)

Starts an application.


## Syntax

_expression_. `AppExecute`( `_Window_`, `_Command_`, `_Minimize_`, `_Activate_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Window_|Optional|**String**|The caption of the application to activate.|
| _Command_|Optional|**String**|The command to start the application. Required if  **Window** is omitted. If the application is running, **Command** is ignored.|
| _Minimize_|Optional|**Boolean**|**True** if the main window is minimized. The default value is **False**.|
| _Activate_|Optional|**Boolean**|**True** if the application is activated. The default value is **True**.|

## Return value

 **Boolean**


## Example

The following example starts and activates Microsoft Excel.


```vb
Sub StartMicrosoftExcel() 
 AppExecute Command:="Excel.exe" 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]