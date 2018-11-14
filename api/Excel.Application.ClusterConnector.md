---
title: Application.ClusterConnector property (Excel)
keywords: vbaxl10.chm133326
f1_keywords:
- vbaxl10.chm133326
ms.prod: excel
api_name:
- Excel.Application.ClusterConnector
ms.assetid: 5382b95a-c796-e638-5c11-5524e4be3acb
ms.date: 06/08/2017
---


# Application.ClusterConnector property (Excel)

Returns or sets the name of the High Performance Computing (HPC) Cluster Connector that is used to run user-defined functions in XLL add-ins. Read/write


## Syntax

 _expression_. `ClusterConnector`

 _expression_ A variable that represents an '[Application](Excel.Application(object).md)' object.


## Return value

 **String**


## Remarks

The setting of the  **ClusterConnector** property corresponds to the **Cluster type** drop-down box under **Formulas** in the **Advanced** category of the **Excel Options** dialog box.




 **Note**  To specify the  **ClusterConnector** property you must install a High Performance Computing (HPC) Cluster Connector. A Cluster Connector enables you to run cluster-safe XLL functions remotely on an HPC cluster for increased performance.

Before you can specify the  **ClusterConnector** property, you must use the **[UseClusterConnector](Excel.Application.UseClusterConnector.md)** property to allow Excel to run user-defined functions in XLL add-ins.


## See also


[Application Object](Excel.Application(object).md)

