---
title: Project.TrackingMethod property (Project)
ms.prod: project-server
api_name:
- Project.Project.TrackingMethod
ms.assetid: cda3f127-5fad-f486-f02d-6d6eeb0d5588
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.TrackingMethod property (Project)

Gets or sets the tracking method used by Project Server for the project. Read/write  **PjProjectServerTrackingMethod**.


## Syntax

_expression_. `TrackingMethod`

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Remarks

The  **TrackingMethod** property is available only in Project Professional, when the project is opened from Project Server. It can be one of the following **[PjProjectServerTrackingMethod](Project.PjProjectServerTrackingMethod.md)** constants: **pjTrackingMethodDefault**, **pjTrackingMethodPercentComplete**, **pjTrackingMethodSpecifyHours**, or **pjTrackingMethodTotalAndRemaining**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]