---
title: TextBox.PostalAddress property (Access)
keywords: vbaac10.chm11050
f1_keywords:
- vbaac10.chm11050
ms.prod: access
api_name:
- Access.TextBox.PostalAddress
ms.assetid: 04fb29c5-909c-a0b8-a4aa-7701abc07037
ms.date: 03/26/2019
localization_priority: Normal
---


# TextBox.PostalAddress property (Access)

You can use the **PostalAddress** property to specify or determine the postal code and the Customer Barcode data corresponding to the address information displayed in a specified field or text box. The PostalAddress Property Wizard enables the setting of these properties. Read/write **String**.


## Syntax

_expression_.**PostalAddress**

_expression_ A variable that represents a **[TextBox](Access.TextBox.md)** object.


## Remarks

For processing the conversion, correct settings are necessary for all properties of a field or text box that will contain a postal code, address, and Customer Barcode data.

For settings, use sections 1 to 3, delimiting with semicolon (;).

### Postal code settings

Specifies the type of postal code for the field or text box.

|Section|Description|
|:-----|:-----|
|1|Specifies the field or text box for Prefecture names.|
|2|Specifies the field or text box for City/Ward/County.|
|3|Specifies the field or text box for Street/Town/Village.|

### Address settings

Specifies that the field or text box contains a postal code or Customer Barcode data.

|Section|Description|
|:-----|:-----|
|1|Specifies the field or text box for postal code.|
|2|Specifies the field or text box for Customer Barcode data.|

> [!NOTE] 
> Two semicolons are required at the end of the value. 

### Customer Barcode data settings

Specifies the type of Customer Barcode data in the field or text box. This setting is the same as the field or text box for postal code.

|Section|Description|
|:-----|:-----|
|1|Specifies the field or text box for Prefecture names.|
|2|Specifies the field or text box for City/Ward/County.|
|3|Specifies the field or text box for Street/Town/Village.|

<br/>

The postal code consists of three address items: Prefecture, City/Ward/County, and Street/Town/Village. Sections in the **PostalAddress** property of a field or text box for a postal code can be omitted. The following table shows how to omit sections from the property setting.

|Property settings|Address input in field or text box|
|:-----|:-----|
|Address1|Address2 \| Address3|
|Address1|Prefecture+City/Ward/County+Street/Town/Village|
|Address1;|Prefecture|
|;Address1|City/Ward/County+Street/Town/Village|
|;Address1;|City/Ward/County|
|;;Address1|Street/Town/Village|
|Address1;Address2|Prefecture \| City/Ward/County+Street/Town/Village|
|Address1;Address1|Prefecture+City/Ward/County+Street/Town/Village|
|Address1;Address2;|Prefecture \| City/Ward/County|
|Address1;Address1;|Prefecture+City/Ward/County|
|;Address1;Address2|City/Ward/County \| Street/Town/Village|
|;Address1;Address1|City/Ward/County+Street/Town/Village|
|Address1;Address2;Address3|Prefecture \| City/Ward/County \| Street/Town/Village|
|Address1;Address2;Address2|Prefecture \| City/Ward/County+Street/Town/Village|
|Address1;Address1;Address2|Prefecture+City/Ward/County \| Street/Town/Village|
|Address1;Address1;Address1|Prefecture+City/Ward/County+Street/Town/Village|


> [!NOTE] 
> The postal code converter program has been developed and licensed by Advanced Giken Corporation for Microsoft Corporation. 



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]