---
title: Project.PhoneticType property (Project)
keywords: vbapj.chm132499
f1_keywords:
- vbapj.chm132499
ms.prod: project-server
api_name:
- Project.Project.PhoneticType
ms.assetid: d959bb6c-9efa-2b4c-594a-1b9294460770
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.PhoneticType property (Project)

Gets or sets the type of characters used to display phonetic information. Read/write  **PjPhoneticType**.


## Syntax

_expression_. `PhoneticType`

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Remarks

The **PhoneticType** property can be one of the following **[PjPhoneticType](Project.PjPhoneticType.md)** constants: **pjKatakanaHalf**, **pjKatakana**, or **pjHiragana**. The **PhoneticType** property produces tangible results only if the Japanese version of Project is used.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]