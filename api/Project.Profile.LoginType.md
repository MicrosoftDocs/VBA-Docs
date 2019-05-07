---
title: Profile.LoginType property (Project)
ms.prod: project-server
api_name:
- Project.Profile.LoginType
ms.assetid: ebf00927-9c84-9fbc-1315-2e95c81c2d68
ms.date: 06/08/2017
localization_priority: Normal
---


# Profile.LoginType property (Project)

Gets or sets the login type for an account profile in Project Professional. Read/write  **[PjLoginType](Project.PjLoginType.md)**.


## Syntax

_expression_. `LoginType`

_expression_ A variable that represents a [Profile](./Project.Profile.md) object.


## Remarks

The  **LoginType** property can be one of the following constants: **pjProjectServerLogin** or **pjWindowsLogin**.


## Example

If the second account profile is a Project Server account, the following statement returns 1, which is the value of the  **pjWindowsLogin** constant.


```vb
Debug.Print Profiles(2).LoginType
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]