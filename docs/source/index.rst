

RUSTRE
######################

Here is the documentation for RUSTRE, a tool for merging and manipulating xlsx files.

RUSTRE can be downloaded for free at: https://github.com/lucsch/rustre

The program currently supports the following operations:

Merging
***********************

Rustre allows to merge several xlsx workbooks with the same structure. The program checks that the
headers are similar before proceeding with the merge

Compare
**********************

Rustre also allows to compare two xlsx files, a source file and a destination file. These two files can have different structures.

The program works as follows: each line of the destination file will be read. A unique key will be created (e.g. a mix of name, first name and date of birth). Depending on the content the following operations will be performed:

- addition: the unique key does not exist in the source file and the line in the destination file must be added.
- modification : the unique key exists in the source file but the content has changed. The source file must be modified to integrate this new content.
- ignore : some lines in the destination file can be ignored.

All these operations are configurable with a configuration file (.ini).

The ini file structure is described in the configuration section.


.. toctree::
   :maxdepth: 4
   :hidden:

   config
   build
   code


