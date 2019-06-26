---
title: Application.CustomizeIMEMode method (Project)
keywords: vbapj.chm254
f1_keywords:
- vbapj.chm254
ms.prod: project-server
api_name:
- Project.Application.CustomizeIMEMode
ms.assetid: 1e6cae3d-7b06-327a-4db1-8b4416d703ee
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.CustomizeIMEMode method (Project)

Customizes which IME mode is used on a given field.


## Syntax

_expression_. `CustomizeIMEMode`( `_FieldID_`, `_IMEMode_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FieldID_|Optional|**Long**|The field to customize. The default value is  **pjTaskName**. Can be one of the **[PjField](Project.PjField.md)** constants|
| _IMEMode_|Optional|**Long**|Specifies the IME mode to use when the focus is on a table column. The default value is  **pjIMEModeNoControl**. Can be one of the **[PjIMEMode](Project.PjIMEMode.md)** constants.|

## Return value

 **Boolean**


## Remarks

The  **CustomizeIMEMode** method produces tangible results only if an East Asian version of Project is used.

Using the  **CustomizeIMEMode** method without specifying any arguments displays the **Customize IME Mode** dialog box.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]