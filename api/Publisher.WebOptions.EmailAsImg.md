---
title: WebOptions.EmailAsImg property (Publisher)
keywords: vbapb10.chm8257545
f1_keywords:
- vbapb10.chm8257545
ms.prod: publisher
api_name:
- Publisher.WebOptions.EmailAsImg
ms.assetid: c44d3b07-2030-4901-b9df-4dcfe08c985c
ms.date: 06/18/2019
localization_priority: Normal
---


# WebOptions.EmailAsImg property (Publisher)

**True** to send the entire publication page as a single JPEG image. Read/write **Boolean**.


## Syntax

_expression_.**EmailAsImg**

_expression_ A variable that represents a **[WebOptions](Publisher.WebOptions.md)** object.


## Return value

Boolean


## Remarks

This property can increase your message's compatibility with older email clients, but may result in larger file size.

This property is accessible for print publications in addition to web publications.

The properties of the **WebOptions** object are used to specify the behavior of web publications. This means that when any of these properties are modified, newly created web publications inherit the modified properties.

This property corresponds to the check box in the **E-Mail Options** section on the **Web** tab of the **Options** dialog box.


## Example

The following example sets Microsoft Publisher to email publication pages as JPEG images.

```vb
Application.WebOptions.EmailAsImg = True
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]