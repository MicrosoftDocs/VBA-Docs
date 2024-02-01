---
title: LabelInfo members (Office)
ms.service: office
ms.assetid: b702522a-fc69-4035-88d3-8075fa713f14
ms.date: 03/29/2021
ms.localizationpriority: medium
description: LabelInfo members (Office)
---

# LabelInfo members (Office)

Represents the label information data object.

The following properties define the **LabelInfo** object. For more information about the concepts of Label metadata, read [Concepts -Label metadata in the MIP SDK](/information-protection/develop/concept-mip-metadata) in the Microsoft Information Protection SDK.

## Properties

|Name|Required/Optional|Readonly|Data type|Description|
|:-----|:-----|:-----|:-----|:-----|
|ActionId|Optional|Yes|String|A lower-case string in xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx format.|
|AssignmentMethod|Required|No|**[MsoAssignmentMethod](msoassignmentmethod-enumeration-office.md)**|Indicates how the label was applied.|
|ContentBits|Optional|Yes|Long|Bitmask that describes the types of content marking that should be applied to a file. CONTENT_HEADER = 0X1, CONTENT_FOOTER = 0X2, WATERMARK = 0X4, ENCRYPT = 0x8|
|IsEnabled|Optional|Yes|Boolean|Indicates whether the classification represented by this label is enabled on the document.|
|Justification|Required|No|String|Text required during label downgrade that justifies the downgraded. |
|LabelId|Required|No|String|LabelId is a unique identifier for each label in an organization. It's returned as a lower-case string in xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx format.|
|LabelName|Required|No|String|Label unique name within the organization. It does not necessarily correspond to display name.|
|SetDate|Optional|Yes|String|Represents timestamp when the label was set. The date string is formatted in Extended ISO 8601 Date Format.|
|SiteId|Optional|Yes|String|SiteId is a unique identifier for each organization in Active Azure Directory. It's returned as a lower-case string in xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx format.|

## Other Properties

|Name|Readonly|Data type|Description|
|:-----|:-----|:-----|:-----|
|Application|Yes|Object|Gets an **Application** object that represents the container application for the **LabelInfo** object.|
|Creator|Yes|Long|Gets a 32-bit integer that indicates the application in which the **LabelInfo** object was created.|


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
