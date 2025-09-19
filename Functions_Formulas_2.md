## Excel Functions Formulas

### TEXT Functions

1. TRIM(), UPPER(), LOWER(), PROPER()
  
2. CONCATENATE
   
   ``` Eg: CONCATENATE("Daniel", " " , "Wright") ```
      =LEFT("Daniel", 3)&" "&("Wright")
      =LEN("Daniel") = 6
   
4. TEXT

   converts a numeric value to text
   =TEXT(value, format_text)
   =“Lisa earned ”&TEXT(B4“$#,###”) returns “Lisa earned $3,725”

5. SEARCH

   SEARCH function returns the number of the character at which a specific character or text string is first found
   SEARCH("%", A11)
   SEARCH("%", A12, 10)
     10 - Starting value

6. IF(ISNUMBER(SEARCH))
   
   =IF(ISNUMBER(SEARCH(“Disp”,A2)),”Display”,”Other”)

   <img width="1333" height="230" alt="image" src="https://github.com/user-attachments/assets/2d71f41a-7080-412b-b116-3b38a433e106" />

### DATE TIME Functions

1. DATEVALUE()
   1/1/1900 = 1
   1/11/1900 = 11
   2/6/2015 12PM = 42041.5
   
2. Fill Series

    Fill days, weekdays months, years

3. TODAY()
     TODAY() = 2/6/2015
     NOW() = 2/6/2015 17:20

4. YEAR, MONTH, DAY< HOUR, MINUE, SECOND

5. EOMONTH()
   Current Date = 8/3/2015

   EOMONTH(C2, 0) = 8/31/2015

    Start of next month
     EOMONTH(C2, 0) + 1 = 9/1/2015
   
7. WEEKDAY()
   
   WEEKDAY(1/1/2024, 1) = 7
   Sun = 0, Mon = 1 ... Sat = 7
   
   
 8. WORKDAY()

    WORKDAY("1/1/2025", 20) = 01/29/2025

    =NETWORKDAYS(S12, S13)

 9. DATEDIF()

       DATEDIF(B2, B3, "D") = 58 days
    
