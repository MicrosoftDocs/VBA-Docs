---
title: ContainerProperties.ContainerStyle property (Visio)
keywords: vis_sdr.chm17651150
f1_keywords:
- vis_sdr.chm17651150
ms.prod: visio
api_name:
- Visio.ContainerProperties.ContainerStyle
ms.assetid: cc7b6757-0287-e25a-9406-554aa70ef181
ms.date: 06/08/2017
localization_priority: Normal
---


# ContainerProperties.ContainerStyle property (Visio)

Determines the visual appearance of the container. Read/write.


## Syntax

_expression_.**ContainerStyle**

 _expression_ An expression that returns a **[ContainerProperties](Visio.ContainerProperties.md)** object.


## Return value

 **Integer**


## Remarks

The value of the **ContainerStyle** property corresponds to the numerical identifier (ID) of the container style that is selected in the **Container Styles** gallery on the **Container Tools Format** tab.

The value of the **ContainerStyle** should always be greater than zero.

If no value is assigned to the **ContainerStyle** property or it is set to a null value, a run-time error ensues. A run-time error also ensues if you assign the property a value that is less than 1 or greater than the maximum ID number in the **Container Styles** gallery.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]