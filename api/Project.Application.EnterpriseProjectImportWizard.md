---
title: Application.EnterpriseProjectImportWizard method (Project)
keywords: vbapj.chm2248
f1_keywords:
- vbapj.chm2248
ms.prod: project-server
api_name:
- Project.Application.EnterpriseProjectImportWizard
ms.assetid: 0666657f-4352-d7d3-5651-88dc584ea917
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.EnterpriseProjectImportWizard method (Project)

Starts the Enterprise  **Project Import Wizard**. Available in Project Professional only.


## Syntax

_expression_. `EnterpriseProjectImportWizard`( `_Name_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Optional|**Variant**|The file name of a single project that is to be imported by using the wizard.|

## Return value

 **Boolean**


## Remarks

Using the  **EnterpriseProjectImportWizard** method still requires that you click **Finish** rather than **Import More** each time the method is used in a macro.

Alternatively, you can use the  **[SaveAs](Project.Project.SaveAs.md)** method in conjunction with the **[EnterpriseResourceGet](Project.Application.EnterpriseResourceGet.md)** method to convert team members into enterprise resources. Using the Map argument in the **SaveAs** method also allows you to specify the import/export map to use when exporting the data.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]