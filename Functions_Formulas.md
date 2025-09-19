## Excel Functions & Formulas

### Tools

![Microsoft Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)

### Functions

#### **Basics**
1. IF
   
   ```Eg: IF(B2<=0,“Yes”,”No”)```

3. NESTED IF

   ```Eg: IF(B2<40,”COLD”,IF(B2>80,”HOT”,”MILD”))```

5. Logical operators within IF
   Here we’re categorizing conditions as “Wet” if the precipitation type equals “rain” OR “snow”, otherwise Conditions = “Dry”

   ```Eg: **IF(OR(F2=“Rain”,F2=“Snow”),“Wet",“Dry")**```

   If the temp is below freezing AND the amount of precipitation > 0, then PrecipType = “Snow”, if the temp is
   above freezing AND the amount of precipitation >0, then PrecipType = “Rain”, otherwise PrecipType = “None”
   Eg: IF(AND(D2=“Yes”,C2>0),“Snow",IF(AND(D2=“No”,C2>0),“Rain",“None"))

6. NOT Operator

   ```Eg: IF(NOT(C2=0),“Wet",“Dry")```

   ```Eg: IF(C2<>0,“Wet",“Dry")```

8. IFERROR(value, value_if_error)

9. Check conditions
  Excel offers a number of different IS formulas, each of which checks whether a certain condition is true:

  ```ISBLANK = Checks whether the reference cell or value is blank```
  
  ```ISNUMBER = Checks whether the reference cell or value is numerical```
  
  ```ISTEXT = Checks whether the reference cell or value is a text string```
  
  ```ISERROR = Checks whether the reference cell or value returns an error```
  
  ```ISEVEN = Checks whether the reference cell or value is even```
  
  ```ISODD = Checks whether the reference cell or value is odd```
  
  ```ISLOGICAL = Checks whether the reference cell or value is a logical operator```
  
  ```ISFORMULA = Checks whether the reference cell or value is a formula```

11. Stats Function
   
  ```=COUNT(A2:A20)```
  
  ```=AVERAGE(A2:A20)```
  
  ```=MEDIAN(A2:A20)```
  
  ```=MODE(A2:A20)```
  
  ```=MAX(A2:A20)```
  
  ```=MIN(A2:A20)```
  
  ```=PERCENTILE(A2:A20,.25)```
  
  ```=PERCENTILE(A2:A20,.75)```
  
  ```=STDEV(A2:A20)```
  
  ```=VAR(A2:A20)```

8. Windows Function

  ```RANK(A2,A2:A8) = 2```
  
  ```RANK(A3,A2:A8) = 7 (lowest)```
  
  ```RANK(A4,A2:A8) = 6```
  
  ```RANK(A5,A2:A8) = 1 (highest)```
  
  ```RANK(A6,A2:A8) = 4```
  
  ```RANK(A7,A2:A8) = 3```
  
  ```RANK(A8,A2:A8) = 5```

10. Find Nth largest/smallest number
  
   ``` LARGE(A2:A8,2) = 90 (the 2nd largest number in the array is 90)```

    ```SMALL(A2:A8,3) = 50 (the 3rd smallest number in the array is 50)```

11. PERCENT RANK

  ``` PERCENTRANK($A$2:$A$19, A14) = 100% (highest)```
  
  ``` PERCENTRANK($A$2:$A$19, A16) = 0% (lowest)```

13. RAND()
    The RAND() function returns a random value between 0 and 1 (to 15 digits)
    The RANDBETWEEN function returns an integer between two values that you specify

#### **Basics 2**
1. SUMPRODUCT
   The SUMPRODUCT formula multiplies corresponding cells from multiple arrays and returns the sum of the products (Note: all arrays must have the same dimensions )

   ```SUMPRODUCT(array1, array2… array_N)```

   ```SUMPRODUCT(B2:B4,C2:C4) = $7.40```

3. SUMPRODUCT
Eg:
    Quantity of goods sold at Shaws:

   ``` SUMPRODUCT((A2:A17=“Shaws”)*C2:C17) = 16 ```
   
    Total revenue from Shaws:

   ``` SUMPRODUCT((A2:A17=“Shaws”)*C2:C17*D2:D17) = $21.80 ```
   
    Revenue from apples sold at Shaws:

   ``` SUMPRODUCT((A2:A17=“Shaws”)*(B2:B17=“Apple”)*C2:C17*D2:D17) = $0.50 ```

5. COUNTIF, SUMIF, AVERAGIF

   Eg: One Condition

   ``` COUNTIF(B2:B20,22) = 2```
   
   ``` SUMIF(A2:A20,“Ryan”,B2:B20) = 190 ```
   
   ``` SUMIF(A2:A20,“<>Tim”,B2:B20) = 702 ```
   
   ``` AVERAGEIF(A2:A20,“Maria”,B2:B20) = 45.75 ```


  Eg: Multiple Conditions
  
   ``` COUNTIFS(B2:B13,“Search”,D2:D13,“>200”) = 3 ```
    
   ``` SUMIFS(D2:D13, A2:A13,“Feb”,B2:B13,“Display”) = 734 ```
    
   ``` AVERAGEIFS(D2:D13, A2:A13,“Jan”,C2:C13,“MSN”) = 263 ```
