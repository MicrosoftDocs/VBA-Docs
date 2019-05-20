---
title: Rules.Save method (Outlook)
keywords: vbaol11.chm2161
f1_keywords:
- vbaol11.chm2161
ms.prod: outlook
api_name:
- Outlook.Rules.Save
ms.assetid: d838eca0-4ec5-ab43-a031-fd65ab7d9f3c
ms.date: 06/08/2017
localization_priority: Normal
---


# Rules.Save method (Outlook)

Saves all rules in the  **[Rules](Outlook.Rules.md)** collection.


## Syntax

_expression_.**Save** (_ShowProgress_)

_expression_ A variable that represents a [Rules](Outlook.Rules.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ShowProgress_|Optional| **Boolean**| **True** to display the progress dialog box, **False** to save rules without showing the progress.|

## Remarks

After you enable a rule, you must also save the rule by using  **Rules.Save** so that the rule and its enabled state will persist beyond the current session. A rule is only enabled after it has been saved successfully.

 **Rules.Save** can be an expensive operation in terms of performance on slow connections to Exchange server. For more information on using the progress dialog box, see [Manage Rules in the Outlook Object Model](../outlook/How-to/Rules/managing-rules-in-the-outlook-object-model.md).

Saving rules that are incompatible or have improperly defined actions or conditions (such as an empty string for  **[TextRuleCondition.Text](Outlook.TextRuleCondition.Text.md)**) will return an error.

The Exchange server limits the maximum number of rules that can be supported by a store.  **Rules.Save** returns an error when this limit is reached.


## See also


[Rules Object](Outlook.Rules.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]