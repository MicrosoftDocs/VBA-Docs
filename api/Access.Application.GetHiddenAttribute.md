---
title: Application.GetHiddenAttribute method (Access)
keywords: vbaac10.chm12570
f1_keywords:
- vbaac10.chm12570
ms.prod: access
api_name:
- Access.Application.GetHiddenAttribute
ms.assetid: aee0e022-08d5-10f8-bfd0-588b5310fb43
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.GetHiddenAttribute method (Access)

The **GetHiddenAttribute** method returns the value of a hidden attribute of a Microsoft Access object in the object's **Properties** dialog box, available by selecting the object in the Database window and choosing **Properties** on the **View** menu.


## Syntax

_expression_.**GetHiddenAttribute** (_ObjectType_, _ObjectName_)

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Required|**[AcObjectType](Access.AcObjectType.md)**|An **AcObjectType** constant that specifies the type of Access object.|
| _ObjectName_|Required|**String**|The name of the Access object.|

## Return value

Boolean


## Remarks

The **GetHiddenAttribute** method (along with the **[SetHiddenAttribute](access.application.sethiddenattribute.md)** method) provides a means of changing an object's hidden attribute from Visual Basic code. With these methods, you can set or read the hidden option available in the object's **Properties** dialog box.

Because the user can set the hidden attributes by selecting or clearing a check box, the **GetHiddenAttribute** method returns **True** if the option setting is **Yes** (the check box is selected) or **False** if the option setting is **No** (the check box is cleared). 

For example, to set an option of this kind by using the **SetHiddenAttribute** method, specify **True** or **False** for the setting argument, as in the following.

```vb
Application.SetHiddenAttribute acTable,"Customers", True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]