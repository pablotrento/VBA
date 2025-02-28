# VBA ToolBOX for Excel 

This VBA ToolBox for Excel was an idea I got after 10 years of working with Excel and VBA along my career, that still on... 

Well basically the issue is that I found myself writing the same functions again and again for different projects so it seemed very logic for me to create this tool-box of functions that made easier to onboard a new solution.

On the other hand, is a way to give back to the community, to share and contribute with colleagues. As many other unknown people in blogs, that I would never probably meet, helped me out eventually to solve some issues in Excel and make my work more valuable. 
Thanks!

So here is this collection... hope you enjoy !

BR
Pablo


# What You will find:

## ArrayFunctions.bas 
### Introduction
Working with arrays in Excel can significantly enhance the efficiency and performance of your VBA code. Arrays allow you to store and manipulate large amounts of data in a structured and efficient manner. 
This repository contains a collection of essential VBA functions for working with arrays in Excel. These functions are designed to simplify common tasks and improve the readability and maintainability of your code.

### Why Use Arrays in Excel?
1. **Efficiency:** Arrays can store and process large amounts of data much faster than using individual cells or ranges.
2. **Memory Management:** Arrays are stored in memory, which can be more efficient than reading and writing to the worksheet.
3. **Flexibility:** Arrays can be easily manipulated using VBA, allowing for complex operations and data transformations.
4. **Readability:** Using arrays can make your code more readable and easier to understand, especially for complex operations.

### List of Functions
#### 1. `ArrayToRange(arr As Variant, rng As Range)`
- **Description:** Transfers the contents of an array to a specified range in the worksheet.
- **Parameters:**
  - `arr`: The array to transfer.
  - `rng`: The range where the array will be placed.
- **Example:**
  ```vba
  Dim myArray(1 To 3, 1 To 2) As Variant
  myArray(1, 1) = "A1"
  myArray(1, 2) = "B1"
  myArray(2, 1) = "A2"
  myArray(2, 2) = "B2"
  myArray(3, 1) = "A3"
  myArray(3, 2) = "B3"
  ArrayToRange myArray, ThisWorkbook.Sheets("Sheet1").Range("A1")
  ```

#### 2. `RangeToArray(rng As Range) As Variant`
- **Description:** Transfers the contents of a specified range in the worksheet to an array.
- **Parameters:**
  - `rng`: The range to transfer.
- **Example if passing a range to the function:**
  ```vba
  Dim myArray As Variant

  We can pass a range as a very basic way
    myArray = RangeToArray(ThisWorkbook.Sheets("Sheet1").Range("A1:B3"))

  Or use CurrentRegion
    myArray = RangeToArray(ws.Range("A1").CurrentRegion)
  
  Select Range from specific Table in specific sheet:
   ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Change "Sheet1" to your sheet name

    ' Set the table
    Set tbl = ws.ListObjects("Table1") ' Change "Table1" to your table name

    myArray = RangeToArray(tbl)

  ```

#### 3. `ArrayResize(arr As Variant, newRows As Long, newCols As Long) As Variant`
- **Description:** Resizes an array to the specified number of rows and columns.
- **Parameters:**
  - `arr`: The array to resize.
  - `newRows`: The new number of rows.
  - `newCols`: The new number of columns.
- **Example:**
  ```vba
  Dim myArray(1 To 2, 1 To 2) As Variant
  myArray(1, 1) = "A1"
  myArray(1, 2) = "B1"
  myArray(2, 1) = "A2"
  myArray(2, 2) = "B2"
  myArray = ArrayResize(myArray, 3, 3)
  ```

#### 4. `ArraySort(arr As Variant, colIndex As Long, ascending As Boolean) As Variant`
- **Description:** Sorts a 2D array based on the values in a specified column.
- **Parameters:**
  - `arr`: The array to sort.
  - `colIndex`: The index of the column to sort by.
  - `ascending`: Whether to sort in ascending (True) or descending (False) order.
- **Example:**
  ```vba
  Dim myArray(1 To 3, 1 To 2) As Variant
  myArray(1, 1) = "B"
  myArray(1, 2) = 2
  myArray(2, 1) = "A"
  myArray(2, 2) = 1
  myArray(3, 1) = "C"
  myArray(3, 2) = 3
  myArray = ArraySort(myArray, 1, True)
  ```

#### 5. `ArrayFilter(arr As Variant, colIndex As Long, criteria As Variant) As Variant`
- **Description:** Filters a 2D array based on a specified column and criteria.
- **Parameters:**
  - `arr`: The array to filter.
  - `colIndex`: The index of the column to filter by.
  - `criteria`: The criteria to filter by.
- **Example:**
  ```vba
  Dim myArray(1 To 3, 1 To 2) As Variant
  myArray(1, 1) = "A"
  myArray(1, 2) = 1
  myArray(2, 1) = "B"
  myArray(2, 2) = 2
  myArray(3, 1) = "A"
  myArray(3, 2) = 3
  Dim filteredArray As Variant
  filteredArray = ArrayFilter(myArray, 1, "A")
  ```

#### 6. `ArrayConcat(arr1 As Variant, arr2 As Variant) As Variant`
- **Description:** Concatenates two arrays.
- **Parameters:**
  - `arr1`: The first array.
  - `arr2`: The second array.
- **Example:**
  ```vba
  Dim arr1(1 To 2, 1 To 2) As Variant
  arr1(1, 1) = "A1"
  arr1(1, 2) = "B1"
  arr1(2, 1) = "A2"
  arr1(2, 2) = "B2"
  Dim arr2(1 To 2, 1 To 2) As Variant
  arr2(1, 1) = "A3"
  arr2(1, 2) = "B3"
  arr2(2, 1) = "A4"
  arr2(2, 2) = "B4"
  Dim concatenatedArray As Variant
  concatenatedArray = ArrayConcat(arr1, arr2)
  ```

#### 7. `ArrayUnique(arr As Variant) As Variant`
- **Description:** Returns an array containing only the unique values from the input array.
- **Parameters:**
  - `arr`: The array to process.
- **Example:**
  ```vba
  Dim myArray(1 To 4) As Variant
  myArray(1) = "A"
  myArray(2) = "B"
  myArray(3) = "A"
  myArray(4) = "C"
  Dim uniqueArray As Variant
  uniqueArray = ArrayUnique(myArray)
  ```

#### 8. `ArrayTranspose(arr As Variant) As Variant`
- **Description:** Transposes a 2D array (rows become columns and columns become rows).
- **Parameters:**
  - `arr`: The array to transpose.
- **Example:**
  ```vba
  Dim myArray(1 To 2, 1 To 3) As Variant
  myArray(1, 1) = "A1"
  myArray(1, 2) = "B1"
  myArray(1, 3) = "C1"
  myArray(2, 1) = "A2"
  myArray(2, 2) = "B2"
  myArray(2, 3) = "C2"
  Dim transposedArray As Variant
  transposedArray = ArrayTranspose(myArray)
  ```


## Functions for Working with Files and Folders

1. **CreateExcelFileFromRange**: Creates a new Excel file from a range of cells.
2. **CreateExcelFileFromArray**: Creates a new Excel file from an array.
3. **CopyFiles**: Copies files from one directory to another.
4. **CreateFolder**: Creates a folder with a name that includes specific date formats.
5. **SaveFileWithDate**: Saves a file with a name that includes specific date formats.


### Examples of Usage

####1. **Create an Excel File from a Range:**

```vba
Sub TestCreateExcelFileFromRange()
    CreateExcelFileFromRange ThisWorkbook.Sheets("Sheet1").Range("A1:B3"), "C:\Users\YourUsername\Documents\NewFile.xlsx"
End Sub
```

####2. **Create an Excel File from an Array:**

```vba
Sub TestCreateExcelFileFromArray()
    Dim myArray(1 To 3, 1 To 2) As Variant
    myArray(1, 1) = "A1"
    myArray(1, 2) = "B1"
    myArray(2, 1) = "A2"
    myArray(2, 2) = "B2"
    myArray(3, 1) = "A3"
    myArray(3, 2) = "B3"
    CreateExcelFileFromArray myArray, "C:\Users\YourUsername\Documents\NewFileFromArray.xlsx"
End Sub
```

####3. **Copy Files from One Directory to Another:**

```vba
Sub TestCopyFiles()
    CopyFiles "C:\Users\YourUsername\Documents\Source", "C:\Users\YourUsername\Documents\Destination"
End Sub
```

####4. **Create a Folder with Date Formats:**

```vba
Sub TestCreateFolder()
    CreateFolder "C:\Users\YourUsername\Documents", "MyFolder", "2025", "01", "-"
End Sub
```

####5. **Save a File with Date Formats:**

```vba
Sub TestSaveFileWithDate()
    SaveFileWithDate "C:\Users\YourUsername\Documents", "MyFile", "yyyy-mm-dd"
End Sub
```







## How to implement Any of the .bas Files with the Macros 

Here are the two basic options for implementing any of the `.bas` files in this REPO... maybe you already know but.. lets review anyways:

### Option 1: Copy and Paste the Code
1. **Open the VBA Editor:**
   - Press `Alt + F11` to open the VBA editor.

2. **Insert a New Module:**
   - In the VBA editor, go to `Insert > Module`.

3. **Paste the Code:**
   - Copy the code for the array functions from your `.bas` file or from the provided code snippet.
   - Paste the code into the new module.

### Option 2: Download and Import the `.bas` File
1. **Download the `.bas` File:**
   - Download the `.bas` file containing the array functions from your repository or source.

2. **Open the VBA Editor:**
   - Press `Alt + F11` to open the VBA editor.

3. **Import the `.bas` File:**
   - In the VBA editor, go to `File > Import File...`.
   - Navigate to the location of the downloaded `.bas` file and select it.
   - Click `Open` to import the file into your VBA project.


### Conclusion
Have funnn! and if you have any more questions or need further assistance, feel free to ask or contact. 🐱‍💻
