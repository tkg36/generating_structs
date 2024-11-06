Struct Constructor Script

Basic Process:

The data is harvested from the sheet(s) by scanning down rows of the A column (name col), searching for where the value of the next is different from the last. The rows that all have the same name in the A column are grouped together
in a Grouping Object (custom object class). With the range of rows defined, all data referring to that struct is grabbed and stored in object variables for easy access. 
This is done all the way down the spreadsheet, returning a list of these Grouping Objects that the following 3 processes use. 
It can be helpful to think of these Grouping Objects as the structs themselves; they contain all the data that the spreadsheet does to describe it.


3 Stages - Making struct definitions, making case statement trees, creating pointers for each struct

Stage 1: Making struct definitions
Iterates through list of these created Grouping Objects, using values stored in the object, to write a struct definition for each to a txt file

Stage 2: Making case statement trees
Iterates through the C column of spreadsheet, searching for any entries that are "same as" another struct
Finds the struct being referred to, adds the two structs to a "family" along with any other struct that is equivalent. 
This is so that there aren't redundant trees, instead multiple equivalent structs share upper level case statements pointing to one tree.
The tree is written. The bulk of the text doesn't change; the upper level case statements are prepended, the title of the "parent" struct is inserted in the tree, and the low level case statements are generated from the "parent" struct and placed within the tree.
*This txt file will likely need to be formatted; the indents may not be consistant. Pasting into Visual Studio seems to correct it into propper C++ formatting.

Stage 3: Creating pointers for each struct
Iterates through the list of Grouping Objects and writes a simple, single line declaring a pointer with the name of the struct.


All stages are done twice, for elemental and nodal structs. Certain global variables are changed to facilitate this, defined in the main method, ('f', 'ws', 'structGroupObjList').
All resulting text is written to a txt file, one for each stage, in the directory the script resides in. 


|--------------------------------------------------------------------------------------------------|



The logic for how the strings are generated can get a little convoluted.
The values of entries in the A-D columns of the spreadsheet can be accessed with simple getter methods of the Grouping Object. The bulk of structs can be handled with just this data. 
There are detector methods for 3 strings of interest, that complicate things:

1) If one of the datatypes specified in the B column is not a primitive data type. This means the datatype is a custom. 
Once detected, the script searches the Typedefs worksheet for the definition of the custom type, which is definied with primatives. Using that definition, the struct definition is constructed.

2) If one of the variable names in the C column is an array. The script detects the presence of '[' or ']'. The script returns a boolean, whether or not it is an array, and a string, the value inside the brackets (blank if not an array).
With these, the struct constructor can then be properly defined, defining each variable as an array of the appropriate size

3) If one of the entries in the C Column contains the words "Same as", meaning the actual struct at this position isn't defined, but is identical to another. Once detected, the script searches the list of Grouping Objects for a struct
with the same title as the struct being pointed to. This should always be successful, barring an error in the spreadsheet. Once found, the constructor function is called recursively to write the definition of the struct being pointed to.
This is why the struct constructor method is split between a header and a body; we want the title of the struct to be the one listed in the sheet, but the body might be of another struct.
*The convention of the "Same as..." entries is not consistant. / and _ are exchanged without reason, and sometimes the full name of the pointed to struct is not defined. There are methods in the script that correct these.
*Should these convention issues be changed in the future, these "fixer" methods shouldn't break anything,



On the topic of the recursion in the makeStructConstructorBody method, there are certain default parameters that aid in keeping data accurate throughout recursion. 
Control is the number of the variable within the struct definition. It gets passed in recursive calls so that the numbering stays consistant, even if a custom variable type is definined in the middle of primatives.
areArrays is a tuple, a bolean and value, that are used to properly define array variables.

The D column of the spreadsheet contains comments. In the interest of making readable and debuggable code, these comments are append (as comments) to the struct definition.




Tyler Grover (tyler.grover@siemens.com, tkg36@drexel.edu)
3/30/23


If you're still confused about some aspect of the script, the code itself is littered with comments. No promises on if they'll make you less confused, but they're there. 
Feel free to contact me for any further questions.



