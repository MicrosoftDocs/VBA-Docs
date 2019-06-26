---
title: Application.MakeFieldEnterprise method (Project)
keywords: vbapj.chm2275
f1_keywords:
- vbapj.chm2275
ms.prod: project-server
api_name:
- Project.Application.MakeFieldEnterprise
ms.assetid: ba9564c9-faa6-bce6-0d59-05dee0cfc887
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.MakeFieldEnterprise method (Project)

Adds a local custom field to Project Server as an enterprise custom field.


## Syntax

_expression_. `MakeFieldEnterprise`( `_FieldID_`, `_FieldName_`, `_LookupTableName_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FieldID_|Required|**Long**|Identification number of the local custom field. Use the  **[FieldNameToFieldConstant](Project.Application.FieldNameToFieldConstant.md)** method to get the FieldID argument.|
| _FieldName_|Required|**String**|Name of the enterprise custom field to create.|
| _LookupTableName_|Optional|**String**|Name of the lookup table to create. The default value is an empty string ("").|

## Return value

 **Boolean**


## Remarks

When the  **MakeFieldEnterprise** method completes successfully, Project shows a dialog box with the message, "The field was successfully added to Project Server. In order to view and use the enterprise field in the project, you will need to quit and restart Project Professional."

The  **MakeFieldEnterprise** method corresponds to the **Add Field to Enterprise** command in the **Custom Fields** dialog box. The method is available only in Project Professional. Project Professional must be connected to Project Server.


## Example

To use the following example, create a local custom field, such as a task text custom field, named  **LocalWithLUT2Enterprise**. Add a lookup table for the custom field that has some values.






```vb
Sub Local2Enterprise() 
 Dim localId As Long 
 localId = FieldNameToFieldConstant(FieldName:="LocalWithLUT2Enterprise") 
 
 MakeFieldEnterprise FieldID:=localId, FieldName:="NewTaskTextFromLocal", LookupTableName:="NewTaskTextLUTFromLocal" 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]