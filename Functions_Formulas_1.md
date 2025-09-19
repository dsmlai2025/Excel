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

### Lookup & Reference Functions

1. VLOOKUP

| A |    B    | C   | D    |    E |
|---|---------|-----|-------|-------|
| 1 | Product | Qty | PID   | Price |
| 2 | T-shirt | 25  | 93764 | =VLOOKUP(A2, $G$1:$H$6,2,0) |

| G        |    H  |
|----------|-----|
| Product  | Price |
| T-shirt  | 10.99  |
| Sweater  | 20.99  |
| Shorts   | 40.99  |
| Socks    | 30.99  |
| Spandex  | 50.99  |

2. HLOOKUP

| A |    B    | C   | D    |    E |
|---|---------|-----|-------|-------|
| 1 | Product | Qty | PID   | Price |
| 2 | T-shirt | 25  | 93764 | =HLOOKUP(A2, $G$1:$J$2,2,0) |

| G        |    H  |  I    |   J   |
|----------|-------|-------|-------|
| Product  |  T-shirt | Sweater  | Shorts |
| Price    | 10.99  |   20.99    | 40.99  |

3. ROW vs ROWS
   
   ROW function returns the row number of a given reference
   ``` Eg: ROW(C10) = 10    ```

   ROWS function returns the number of rows in a given arrayor

   ``` ROWS(A10:D15) = 6 ```

   ``` ROWS({1,2,3;4,5,6}) = 2    ```
   Returns number of rows of an array
   
5. COLUMN vs COLUMNS

  ``` COLUMN(C10) = 3   ```
  
  ``` COLUMNS(A10:D15) = 4   ```
  
  ``` COLUMNS({1,2,3;4,5,6}) = 3   ```

6. INDEX

  Eg: Returns the value R5, C3
   
  ``` INDEX($A$1:$C$5, 5, 3) = 234 ```

7. MATCH

   The MATCH function returns the position of a specific value within a column or row  
   Eg: MATCH(lookup_value, lookup_array, [match_type])

   match_type = 0 (Exact match)
   match_type = 1 (<= Lookup value)
   match_type = -1 (>= Lookup value)
   
    |    | A        |    B  |
    |----|----------|-----|
    | 1   | Tools  | Price |
    | 2   | Hammer  | 10.99  |
    | 3   | Screw Driver  | 20.99  |
    | 4   | Pliers   | 40.99  |
    | 5   | Wrench    | 30.99  |

   MATCH(“Pliers”,$A$1:$A$5, 0) = 4

   <img width="1027" height="222" alt="image" src="https://github.com/user-attachments/assets/9fd15e2c-6e51-48cf-84b3-617ec940346a" />


8. INDEX and MATCH

<img width="800" height="600" alt="image" src="https://github.com/user-attachments/assets/65d9e176-38cb-43ea-945f-c561e9bab697" />

9. XLOOKUP

   XLOOKUP replaces older functions like VLOOKUP and HLOOKUP by being more flexible, allowing both vertical and horizontal lookups. can retrieve a dynamic array of results.
   
   =XLOOKUP("T-shirt", A2:A4, B2:B4, "Not Found")
   
      | A        |    B  |
      |----------|-----|
      | Product  | Price |
      | T-shirt  | 10.99  |
      | Sweater  | 20.99  |
      | Shorts   | 40.99  |

11. CHOOSE
    
      The CHOOSE function selects a value, cell reference, or function to perform from a list, based on a given index number

    ``` =CHOOSE(3, "Apple", "Banana", "Cherry", "Date") ```
    Returns Cherry
    
    ``` =CHOOSE(2, 100, 200, 300) ```
    Return 200

    Assuming the following values are in cells A1=10, A2=20, A3=30:
    ``` =CHOOSE(1, A1, A2, A3) ```
    Return 10
    
13. OFFSET
    
      return either the value of a cell within an array (like INDEX) or a specific range of cells

      =OFFSET(reference, rows, columns, [height], [width])

    Eg: ``` =OFFSET(A1, 3, 1) ```
   
    Eg 2:

    ``` =OFFSET(A7, 0, 2, 5, 1) ```
    <img width="200" height="200" alt="image" src="https://github.com/user-attachments/assets/67111e0e-8222-42cd-9eb4-1be57296dc51" />

    <img width="119" height="88" alt="image" src="https://github.com/user-attachments/assets/7a3bd798-1753-4ff9-b079-02ca0cb4f880" />

    Eg 3: 
   
   Assume E1 = 3
  
   ``` =SUM(OFFSET(B1, COUNT(B:B)-E1, 0, E1, 1)) ```
   
   Expl: SUM(OFFSET(B1, 3, 0, 3, 1)) = SUM(B4:B6) = 20 + 25 + 30 = 75
         |  B |
         |----|
         | 5  |
         | 10 |
         | 15 |
         | 20 |
         | 25 |
         | 30 |
         
   
