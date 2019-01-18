---
title: Project.AskForCompletedWork Property (Project)
ms.prod: project-server
api_name:
- Project.Project.AskForCompletedWork
ms.assetid: 54380c01-ae6f-a378-a46b-bfe0064fbc5f
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.AskForCompletedWork Property (Project)

Gets or sets the way completed work is reported in team status messages. Read/write  **PjTeamStatusCompletedWork**.


## Syntax

 _expression_. `AskForCompletedWork`

 _expression_ A variable that represents a [Project](./Project(enumerations).md) object.


## Remarks

The  **AskForCompletedWork** property can be one of the following **[PjTeamStatusCompletedWork](Project.PjTeamStatusCompletedWork.md)** constants: **pjBrokenDownByDay**, **pjBrokenDownByWeek**, or **pjTotalForEntirePeriod**.


