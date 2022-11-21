
# Calculating actual vehicle revenue miles & hours

This code is basically analyzing *National Transit Database (NTD)* to calculate actual vehicle revenue miles & hours and actual total vehicle miles & hours. 
The **purpose** of developing this code is to generate two files which are needed to report to Federal Transit Administration, i.e., MR-20 and S-10.



## File Format

This project is used by the following file formats:

- **seperate 10 .xlsx files**, which is extracted from *LeeTran NTD Workbook*


## Deployment

To deploy this project run, the following modules are needed to be imported as belows.

```bash
import os
import pandas as pd
import numpy as np
import datetime
from datetime import date, timedelta
```



## Repository Structure

#### Update key notes:


- (1) Modify VOMS, service changes and atypical days
- (2) Create No.14 Service Chnages VOMS
- (3) Automatically generate S-10 and verified


| File Name | Type     | Description                |
| :-------- | :------- | :------------------------- |
| `NTD_MB_11_18_2022` | `.py` | **Required**.  |

#### Other supplementary files description


| File Name | Type     | Description                       |
| :-------- | :------- | :-------------------------------- |
| `Daily Ridership by Route`      | `.xlsx` | **input** |
| `Service changes within the RY/FY`      | `.xlsx` | **input** RY: reporting year / FY: fiscal year|
| `Scheduled VRM & VRH`      | `.xlsx` | **input**  vehicle revenue miles & vehicle revenue hours|
| `Scheduled DHM & DHH`      | `.xlsx` | **input**  deadhead miles & deadhead hours|
| `Atypical days`      | `.xlsx` | **Deviation Tables (input)** |
| `Added Runs`      | `.xlsx` | **Deviation Tables (input)** |
| `Lost Runs`      | `.xlsx` | **Deviation Tables (input)** |
| `VOMS`      | `.xlsx` | **input** Vehicles Operated at Maximum Service (the number of VOMS changes dependent on service change|
| `Actual VRM & VRH`      | `.xlsx` | **output** actual vehicle revenue miles & actual vehicle revenue hours|
| `Actual TVM & TVR`      | `.xlsx` | **output** actual total vehicle miles & actual total vehicle hours|
| `MR-20`      | `.xlsx` | **output** Automated Monthly Form (needed to report to Federal Transit Administration)|
| `S-10`      | `.xlsx` | **output** Automated Annually Form (needed to report to Federal Transit Administration)|



## Author

- [@xw0413happy](https://github.com/xw0413happy)


## 🚀 About Me
I took 2 python classes during my M.S. degree-seeking program (Civil Engineering), now I am a computer language amateur, strong desire to learn more.


## 🛠 Skills
Python, R, SQL, ArcGIS, Nlogit, Stata, Power BI, Javascript, HTML, CSS, Synchro, Vissim, AutoCAD, Tableau, VBA


## Acknowledgements

 - [FTA](https://www.transit.dot.gov/)
 - [Learn more about NTD](https://www.transit.dot.gov/ntd)

