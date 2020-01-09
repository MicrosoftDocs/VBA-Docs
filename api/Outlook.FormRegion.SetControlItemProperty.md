---
title: FormRegion.SetControlItemProperty method (Outlook)
keywords: vbaol11.chm2401
f1_keywords:
- vbaol11.chm2401
ms.prod: outlook
api_name:
- Outlook.FormRegion.SetControlItemProperty
ms.assetid: da0b3762-e10d-85d1-70bf-94156d21e900
ms.date: 06/08/2017
localization_priority: Normal
---


# FormRegion.SetControlItemProperty method (Outlook)

Binds an explicit built-in property or a custom property to a control in the form region.


## Syntax

_expression_. `SetControlItemProperty`( `_Control_` , `_PropertyName_` )

_expression_ A variable that represents a [FormRegion](Outlook.FormRegion.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Control_|Required| **Object**|A control in the form region to which the property specified by  _PropertyName_ will be bound.|
| _PropertyName_|Required| **String**|The name of the property that will be bound to the control.|

## Remarks

You can use this method to bind an explicit built-in property or a custom property to a control. You must reference the property by its string name, for example, **Subject**, and not by namespace, for example, `http://schemas.microsoft.com/mapi/proptag/0x0037001E`.

The  _PropertyName_ parameter is not case-sensitive. For example, **SetControlItemProperty** interprets an argument, _CustomerId_, to be the same as  _CustomerID_ and binds the specified control to the built-in **[ContactItem.CustomerID](Outlook.ContactItem.CustomerID.md)** property.

Other than using the **SetControlItemProperty** method of a **[FormRegion](Outlook.FormRegion.md)** object, you can also use code similar to the following to bind a property such as the **Subject** property to a control:




```vb
myPage.Controls("bar").ItemProperty = "subject"
```


## See also


[FormRegion Object](Outlook.FormRegion.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
