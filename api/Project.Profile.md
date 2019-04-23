---
title: Profile object (Project)
ms.prod: project-server
api_name:
- Project.Profile
ms.assetid: 92ae9d1a-ea4d-1814-1655-f0798f4b18d0
ms.date: 06/08/2017
localization_priority: Normal
---


# Profile object (Project)


 

Represents an account profile in Project Professional. The  **Profile** object is a member of the **[Profiles](Project.profiles.md)** collection.
 
If the second account profile is a Project Server account, the following statement returns the value 1, which is the value of the  **pjServerProfile** constant in the **[PjProfileType](Project.PjProfileType.md)** enumeration.
 



```vb
Debug.Print Profiles(2).Type
```


## Remarks

The  **Project Server Accounts** dialog box shows the number and order of profiles. Use `Profiles.Count` to programmatically determine the number of account profiles defined in Project Professional.
 

 

## Methods



|Name|
|:-----|
|[Delete](Project.Profile.Delete.md)|

## Properties



|Name|
|:-----|
|[ConnectionState](Project.Profile.ConnectionState.md)|
|[LoginType](Project.Profile.LoginType.md)|
|[Name](Project.Profile.Name.md)|
|[Server](Project.Profile.Server.md)|
|[SiteId](Project.profile.siteid.md)|
|[Type](Project.Profile.Type.md)|
|[UserName](Project.Profile.UserName.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]