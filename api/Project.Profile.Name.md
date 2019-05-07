---
title: Profile.Name property (Project)
ms.prod: project-server
api_name:
- Project.Profile.Name
ms.assetid: 98e1ca12-ecaa-aaae-de48-352301c28e50
ms.date: 06/08/2017
localization_priority: Normal
---


# Profile.Name property (Project)

Gets the name of an account profile in Project Professional. Read/write  **String**.


## Syntax

_expression_.**Name**

_expression_ A variable that represents a [Profile](./Project.Profile.md) object.


## Remarks

The  **Project Server Accounts** dialog box shows the number and order of profiles. Use `Profiles.Count` to programmatically determine the number of account profiles.


## Example

If the second account profile is a Project Server account, the following statement returns the name of the account.


```vb
Debug.Print Profiles(2).Name
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]