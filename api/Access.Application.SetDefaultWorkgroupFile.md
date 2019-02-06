---
title: Application.SetDefaultWorkgroupFile method (Access)
keywords: vbaac10.chm12595
f1_keywords:
- vbaac10.chm12595
ms.prod: access
api_name:
- Access.Application.SetDefaultWorkgroupFile
ms.assetid: 64dc24a0-e6dc-685f-620a-463417e8a25d
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.SetDefaultWorkgroupFile method (Access)

Sets the default workgroup file to the specified file.


## Syntax

_expression_.**SetDefaultWorkgroupFile** (_Path_)

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Path_|Required|**String**|The full path and file name of the workgroup file to use as the default.|

## Return value

Nothing


## Remarks

If the file specified by _Path_ does not exist, an error occurs.


## Example

The following example sets the default workgroup file to the file system.mdw in the directory C:\Documents and Settings\Wendy Vasse\Application Data\Microsoft\Access.


```vb
Application.SetDefaultWorkgroupFile _ 
 Path:="C:\Documents and Settings\Wendy Vasse\" _ 
 & "Application Data\Microsoft\Access\system.mdw"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]