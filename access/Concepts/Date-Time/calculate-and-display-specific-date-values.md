---
title: Calculate and display specific date values
description: Create expressions and custom functions for displaying specific dates and calculating time intervals.
ms.prod: access
ms.assetid: ba8c8404-fbe9-d7ef-57bb-17631ec8fb4c
ms.date: 09/21/2018
localization_priority: Normal
---


# Calculate and display specific date values

Because a Date value is stored as a double-precision number, you may receive incorrect formatting results when you try to manipulate Date values in an expression. This topic illustrates how to create expressions and custom functions for displaying specific dates and calculating time intervals. 


**Current month** 

```vb
DateSerial(Year(Date()), Month(Date()), 1)
```

**Next month**

```vb
DateSerial(Year(Date()), Month(Date()) + 1, 1)
```

**Last day of the current month** 

```vb
DateSerial(Year(Date()), Month(Date()) + 1, 0)
```

**Last day of the next month** 

```vb
DateSerial(Year(Date()), Month(Date()) + 2, 0)
```

**First day of the previous month** 

```vb
DateSerial(Year(Date()), Month(Date())-1,1)
```

**Last day of the previous month**

```vb
DateSerial(Year(Date()), Month(Date()),0)
```

**First day of the current quarter**

```vb
DateSerial(Year(Date()), Int((Month(Date()) - 1) / 3) * 3 + 1, 1)
```

**Last day of the current quarter**

```vb
DateSerial(Year(Date()), Int((Month(Date()) - 1) / 3) * 3 + 4, 0)
```

**First day of the current week (assuming Sunday = day 1)**

```vb
Date() - WeekDay(Date()) + 1
```

**Last day of the current week**

```vb
Date() - WeekDay(Date()) + 7
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]