
## Functions and Formulas

### 1. Dynamic Excel
  
Formulas that can return arrays of variable size are called dynamic arrays; these
formulas are entered in a single cell and can “spill” results across an entire range

Product	Sales	Margin	Profit
Smart Speaker	14891	20%	2978.2
Bat	3927	5%	196.35
Vinyl Records	2447	8%	195.76
Lego	8510	6%	510.6
Sunglasses	6731	2%	134.62
Blu Ray Player	2744	3%	82.32

<img width="400" height="150" alt="image" src="https://github.com/user-attachments/assets/22b36c69-f2df-4eec-aeef-5617eac97ae5" />

### 2. UNIQUE()

  

### 3. SORT()

=SORT( array, [sort_index], [sort_order], [by_col] )
=SORT(A2:D10, 4, -1) 
  Sorted by 4th col DESC
  
=SORT(A2:D10, {3,4}, {1, -1} )
  Sorted by 3rd col, ASC, 4th col DESC

### 4. FILTER()

  Returns matching records.
  
  =FILTER(A2:C10, B2:B10, "No Results")
  
    <img width="400" height="400" alt="image" src="https://github.com/user-attachments/assets/b56fdab6-8d2a-45f6-a29f-1a48c7cc4a81" />
  
  FILTER AND
  
    <img width="400" height="400" alt="image" src="https://github.com/user-attachments/assets/c62e2c5e-1170-4f3c-b56f-c24677a9436e" />
  
  FILTER OR
  
  <img width="1054" height="297" alt="image" src="https://github.com/user-attachments/assets/05245aac-4197-4da8-8f13-8ddaadeb16d2" />

### 5. UNIQUE()

    Returns unique values

    UNIQUE(B2:B10)

### 6. SORT and UNIQUE
  
    SORT(UNIQUE(B2:B10))

### 7. SORT and FILTER
    
  <img width="1054" height="284" alt="image" src="https://github.com/user-attachments/assets/a97b9863-7c25-4ac5-8f42-eecbe0febd96" />

### 8. SEQUENCE
     
<img width="500" height="300" alt="image" src="https://github.com/user-attachments/assets/cbb9ac73-7271-4cfd-9690-30bf390831d7" />

### 8. RANDARRAY

    =RANDARRAY( [rows], [columns], [min], [max], [integer] )
    
### 9. FREQUENCEY

    =FREQUENCY( data_array, bins_array )
    
  <img width="600" height="300" alt="image" src="https://github.com/user-attachments/assets/9b46dd4c-faf4-4aa1-b6ad-008ac99ba54d" />

### 10. TRANSPOSE

   <img width="202" height="127" alt="image" src="https://github.com/user-attachments/assets/d71b69bb-7985-4c0a-bf14-d30293be53d3" />
   <img width="410" height="52" alt="image" src="https://github.com/user-attachments/assets/24aeb666-b045-417c-9863-3bf97ff8cf57" />

### 10. LET

   <img width="600" height="271" alt="image" src="https://github.com/user-attachments/assets/ded80a8c-e808-40ab-9a1b-3e05a75ace59" />

### 11. INDIRECT
    
  <img width="341" height="169" alt="image" src="https://github.com/user-attachments/assets/bb1f515b-8fa1-4511-a5b6-aaf57434062b" />

    SUM(INDIRECT(D2)) = 16

### 11. HYPERLINK  

    =HYPERLINK(”http://www.example.com”, “Click Here”)
