
How to identify duplicates in Excel
Input the above formula in B2, then select B2 and drag the fill handle to copy the formula down to other cells:
=IF(COUNTIF($A$2:$A$8, $A2)>1, "Duplicate", "Unique")
The formula will return "Duplicates" for duplicate records, and a blank cell for unique records:
Mar 21, 2023


https://www.ablebits.com/office-addins-blog/identify-duplicates-excel/
