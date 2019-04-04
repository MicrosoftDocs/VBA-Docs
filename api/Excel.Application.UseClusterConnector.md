---
title: Application.UseClusterConnector property (Excel)
keywords: vbaxl10.chm133325
f1_keywords:
- vbaxl10.chm133325
ms.prod: excel
api_name:
- Excel.Application.UseClusterConnector
ms.assetid: 9da42299-f23d-66e8-79b3-6105a0626db1
ms.date: 04/05/2019
localization_priority: Normal
---


# Application.UseClusterConnector property (Excel)

Returns or sets whether Excel allows user-defined functions in XLL add-ins to be run on a compute cluster. Read/write.


## Syntax

_expression_.**UseClusterConnector**

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Return value

**Boolean**


## Remarks

**True** if Excel allows user-defined functions in XLL add-ins to be run on a compute cluster; otherwise, **False**.

The setting of the **UseClusterConnector** property corresponds to the **Allow user-defined XLL functions to run on a compute cluster** check box under **Formulas** in the **Advanced** category of the **Excel Options** dialog box.

> [!NOTE] 
> To enable the **UseClusterConnector** property, you must install a High Performance Computing (HPC) Cluster Connector. A Cluster Connector enables you to run cluster-safe XLL functions remotely on an HPC cluster for increased performance.

After setting the **UseClusterConnector** property, use the **[ClusterConnector](Excel.Application.ClusterConnector.md)** property to specify the HPC Cluster Connector to use.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]