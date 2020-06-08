Attribute VB_Name = "Data_Module"
Option Explicit
'Data_Module – implements intraworkbook data storage and manipulation
'   Schema:
'       Cache
'           (Metadata)  First four cells are for cache metadata
'                           (1, 1) - column name
'                           (2, 1) - block name
'                           (2, 2) - block size
'           (Data)      First two rows (beginning after metadata, then ongoing)
'                           (1, j+0) - key
'                           (1, j+1) - value
'       Storage
'           (Metadata)  First two cells are for column metadata
'                           (3, j+0) - Column Name
'                           (3, j+1) - block size
'           (Data)      Every two columns are for storing data columns
'                           (i+0, j+0) - block name
'                           (i+n, j+0) - keys   (n = block size)
'                           (i+n, j+1) - values (n = block size)
'
'

Public Sub set_data_s(col_name As String, block_name As String, value As Variant)
'Sets singleton data block
'   given column name, block name, and value
'   In a singleton data block, block name is used as key

Dim result As Integer

'Print cache metadata
Data_Sheet.Cells(1, 1) = col_name
Data_Sheet.Cells(2, 1) = block_name
Data_Sheet.Cells(2, 2) = 0

'Print cache data
Data_Sheet.Cells(1, 3) = value

'Write cache block
result = write_block

'Check for errors
If result = 1 Then
    MsgBox "ERROR (set_data): mismatched block sizes in column " & col_name
    Exit Sub
End If

End Sub
Public Function get_data_s(col_name As String, block_name As String)
'Gets singleton data block
'   given column name, block name, value
'   In a singleton data block, block name is used as key

Dim result As Integer

result = cache_block(col_name, block_name)

If result = 1 Then
    MsgBox "ERROR (get_data): column " & col_name & " not found"
    Exit Function
ElseIf result = 2 Then
    MsgBox "ERROR (get_data): block " & block_name & " not found"
    Exit Function
End If

get_data_s = Data_Sheet.Cells(1, 3)

End Function

Public Sub set_data_m(col_name As String, block_name As String, key As String, value As Variant)
'Sets value in single key:value pair of multi data block
'   given a column name, block name, key, and value

Dim result As Integer

result = cache_block(col_name, block_name)

If result = 1 Then
    MsgBox "ERROR (update_data): column " & col_name & " not found"
ElseIf result = 2 Then
    MsgBox "ERROR (update_data): block " & block_name & " not found"
End If

result = search_cache(key)

If result = -1 Then
    MsgBox "ERROR (update_data): key " & key & " not found"
    Exit Sub
End If

Data_Sheet.Cells(2, result) = value

result = write_block

If result = 1 Then
    MsgBox "ERROR (set_data_m): mismatched block sizes in column " & col_name
    Exit Sub
End If

End Sub

Public Function get_data_m(col_name As String, block_name As String, key As String) As Variant
'Gets value in single key:value pair of multi data block
'   given a column name, block name, key, and value

Dim result As Integer

result = cache_block(col_name, block_name)

If result = 1 Then
    MsgBox "ERROR (get_data): column " & col_name & " not found"
    Exit Function
ElseIf result = 2 Then
    MsgBox "ERROR (get_data): block " & block_name & " not found"
    Exit Function
End If

result = search_cache(key)

If result = -1 Then
    MsgBox "ERROR (get_data): key " & key & " not found"
    Exit Function
End If

get_data_m = Data_Sheet.Cells(2, result)

End Function

Public Function search_cache(key As String) As Variant

Dim j As Integer

j = 3

Do Until Data_Sheet.Cells(1, j) = "" Or Data_Sheet.Cells(1, j) = key
    j = j + 1
Loop

If Data_Sheet.Cells(1, j) = "" Then
    search_cache = -1
Else
    search_cache = j
End If

End Function

Public Function cache_block(col_name As String, block_name As String) As Integer
'Given column name and block name:
'   Retrieves block from data and stores it in cache

Dim i As Integer, j As Integer, k As Integer
Dim block_size As Integer

j = 1   'designates column search to begin with 1st column
i = 4   'designates block search to begin with 4th row
        '   (rows 1,2 are cache, row 3 is col metadata)

'Loops until correct column is found, or until there are no more columns
Do Until Data_Sheet.Cells(3, j) = col_name Or Data_Sheet.Cells(3, j) = ""
    j = j + 2
Loop

'If empty, return empty
If Data_Sheet.Cells(3, j) = "" Then
    cache_block = 1
    Exit Function
End If

'gets block_size from where it is listed (next to col_name)
block_size = Data_Sheet.Cells(3, j + 1)

'Loops until correct block is found, or until there are no more blocks
Do Until Data_Sheet.Cells(i, j) = block_name Or Data_Sheet.Cells(i, j) = ""
    i = i + block_size + 1
Loop

'If empty, return empty
If Data_Sheet.Cells(i, j) = "" Then
    cache_block = 2
    Exit Function
End If

'Prints block to cache
If block_size = 0 Then
    'Single Item Block
    Data_Sheet.Cells(1, 3) = Data_Sheet.Cells(i, j + 1)
Else
    'Multi Item Block
    For k = 0 To block_size - 1
        Data_Sheet.Cells(1, 3 + k) = Data_Sheet.Cells(i + k + 1, j)
        Data_Sheet.Cells(2, 3 + k) = Data_Sheet.Cells(i + k + 1, j + 1)
    Next
End If

'Print cache metadata
Data_Sheet.Cells(1, 1) = col_name
Data_Sheet.Cells(2, 1) = block_name
Data_Sheet.Cells(2, 2) = block_size

cache_block = 0

End Function

Public Function write_block() As Integer
'Given that a block is written in cache:
'   Writes block to data

Dim i As Integer, j As Integer, k As Integer
Dim col_name As String
Dim block_name As String
Dim block_size As Integer

col_name = Data_Sheet.Cells(1, 1)
block_name = Data_Sheet.Cells(2, 1)
block_size = Data_Sheet.Cells(2, 2)

j = 1   'designates column search to begin with 1st column
i = 4   'designates block search to begin with 4th row
        '   (rows 1,2 are cache, row 3 is col metadata)

'Loops until correct column is found, or until there are no more columns
Do Until Data_Sheet.Cells(3, j) = col_name Or Data_Sheet.Cells(3, j) = ""
    j = j + 2
Loop

'If empty, make new column
If Data_Sheet.Cells(3, j) = "" Then
    Data_Sheet.Cells(3, j) = col_name
    Data_Sheet.Cells(3, j + 1) = block_size
End If

'checks block_size with where it is listed (next to col_name)
If Data_Sheet.Cells(3, j + 1) <> block_size Then
    write_block = 1
    Exit Function
End If

'Loops until correct block is found, or until there are no more blocks
Do Until Data_Sheet.Cells(i, j) = block_name Or Data_Sheet.Cells(i, j) = ""
    i = i + block_size + 1
Loop

'If empty, make new block
If Data_Sheet.Cells(i, j) = "" Then
    Data_Sheet.Cells(i, j) = block_name
End If

'Prints cache to block, clears cache
If block_size = 0 Then
    'Single Item Block
    Data_Sheet.Cells(i, j + 1) = Data_Sheet.Cells(1, 3)
    Data_Sheet.Cells(1, 3) = ""
Else
    'Multi Item Block
    For k = 0 To block_size - 1
        Data_Sheet.Cells(i + k + 1, j) = Data_Sheet.Cells(1, 3 + k)
        Data_Sheet.Cells(1, 3 + k) = ""
        Data_Sheet.Cells(i + k + 1, j + 1) = Data_Sheet.Cells(2, 3 + k)
        Data_Sheet.Cells(2, 3 + k) = ""
    Next
End If

'Clear cache metadata
Data_Sheet.Cells(1, 1) = ""
Data_Sheet.Cells(2, 1) = ""
Data_Sheet.Cells(2, 2) = ""

write_block = 0

End Function

Public Sub generate_block(col_name As String, block_name As String, block_size As Integer)

Dim k As Integer

'Print cache metadata
Data_Sheet.Cells(1, 1) = col_name
Data_Sheet.Cells(2, 1) = block_name
Data_Sheet.Cells(2, 2) = block_size

'Creates block to cache
For k = 0 To block_size - 1
    Data_Sheet.Cells(1, 3 + k) = "Data" & k
    Data_Sheet.Cells(2, 3 + k) = WorksheetFunction.RandBetween(1, 1000)
Next

End Sub

Public Sub test()

    Call generate_block("random_numbers", "A", 10)
    Call write_block
    Call generate_block("random_numbers", "B", 10)
    Call write_block
    Call generate_block("random_numbers", "A", 10)
    Call write_block
    Call generate_block("random_numbers", "C", 10)
    Call write_block
    Call generate_block("other_random_numbers", "A", 4)
    Call write_block
    Call generate_block("other_random_numbers", "B", 4)
    Call write_block
    
    Call set_data_s("new_col", "A", 1)
    Call set_data_s("new_col", "B", 1)
    Call set_data_s("new_col", "A", 2)
    
    MsgBox get_data_s("new_col", "A")
    
    Call set_data_m("other_random_numbers", "A", "Data0", "poop")
    
    MsgBox get_data_m("other_random_numbers", "A", "Data0")

    Call cache_block("random_numbers", "A")

End Sub





