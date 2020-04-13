---
title: Profile.ConnectionState property (Project)
keywords: vbapj.chm131665
f1_keywords:
- vbapj.chm131665
ms.prod: project-server
api_name:
- Project.Profile.ConnectionState
ms.assetid: df961e3e-26a2-9b70-475d-143b2a6db7cb
ms.date: 06/08/2017
localization_priority: Normal
---


# Profile.ConnectionState property (Project)

Gets the connection state of Project Professional, which allows you to determine whether the online mode is for a local profile or for Project Server. Read-only  **PjProfileConnectionState**.


## Syntax

_expression_. `ConnectionState`

_expression_ A variable that represents a [Profile](./Project.Profile.md) object.


## Remarks

The **ConnectionState** property can be one of the following **[PjProfileConnectionState](Project.PjProfileConnectionState.md)** constants: **pjProfileOffline** or **pjProfileOnline**.

You can use this property in conjunction with the **Profile**. **[Type](Project.Profile.Type.md)** property to determine whether the online mode is for a local profile or for Project Server. This property is available only in Project Professional.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]