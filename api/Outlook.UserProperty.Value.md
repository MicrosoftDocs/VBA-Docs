---
title: UserProperty.Value property (Outlook)
keywords: vbaol11.chm222
f1_keywords:
- vbaol11.chm222
ms.prod: outlook
api_name:
- Outlook.UserProperty.Value
ms.assetid: 9f313262-ffd4-3245-f516-bc2d62d6f33a
ms.date: 06/08/2017
localization_priority: Normal
---


# UserProperty.Value property (Outlook)

Returns or sets a  **Variant** indicating the value for the specified custom property. Read/write.


## Syntax

_expression_.**Value**

_expression_ A variable that represents a [UserProperty](Outlook.UserProperty.md) object.


## Remarks

To set for the first time a property created by the  **[UserProperties.Add](Outlook.UserProperties.Add.md)** method, use the **UserProperty.Value** property instead of the **[SetProperties](Outlook.PropertyAccessor.SetProperties.md)** or **[SetProperty](Outlook.PropertyAccessor.SetProperty.md)** method of the **[PropertyAccessor](Outlook.PropertyAccessor.md)** object.

For more information on accessing properties in Outlook, see [Properties Overview](../outlook/How-to/Navigation/properties-overview.md).


## See also


[UserProperty Object](Outlook.UserProperty.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]