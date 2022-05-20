# Configuration file

To work, "RUSTRE" needs a configuration file (*.ini file) describing the data model. 
The structure of this file is explained below

## Groups
Two identical groups must be present:

- source `[SOURCE]` describe the xlsx source file structure. 
- target `[TARGET]` describe the xlsx target file structure.

:::{Note}  
The source file is the reference. Only this file can be modified according to the values found in the target file
:::

## Values

For each group, the following values are required:

| Option         | Value  | Description                                                                                               | Exemple               |
|----------------|--------|-----------------------------------------------------------------------------------------------------------|-----------------------|
| id_col         | list   | the columns used to create an identity (column index is zero based)                                       | `id_col = 0,1,3`        |
| skip_col       | number | the index of the column in which to check if the data should be skipped (Zero based index, could be None) | `skip_col = 7`          |
| skip_col_value | text   | if the parsed cell (in column skip_col) is equal to this value, we ignore the row when parsing the file   | `skip_col_value = Test` |
| col_compare    | number | The comparison column between source and target files                                                     | `col_compare = 4`       |
| col_order      | list   | The order of the columns so that those in the target file match those in the source file                  | `col_order = 0,1,7,8`   |


## Sample


An example of a complete configuration file is available bellow (see also `test/data/test_compare.ini`)

```ini
[SOURCE]
id_col= 0,1,3
skip_col =
skip_col_value =
col_compare = 8
col_order = 0,1,2,3,4,5,6,7,8

[TARGET]
id_col = 0,1,3
skip_col = 7
skip_col_value = Chef / Boss
col_compare = 4
col_order = 0,1,2,3,5,6,7,8,4
```






