README
======

Requirements
============

    - Python 2.7 
    - xlrd
    - mwclient

How to use
==========

    python2 semanticParser [options] username:password excel_file
    
    Examples:
    python2 semanticParser --upload username:password sample.xls
    python2 semanticParser --files sample.xls

Options
=======
    
    -u --upload     
        Uploads content directly to the wiki by creating new pages.

    -f --files
        Generates outputfiles and stores them in the working directory.

Excel format
============


    You should have recieved a copy of sample.xls along with this software.
    
    The format needs to be as follows:

    Row[1]:     propertie description
    Row[2]:     effective propertie names
    Row[3]:     data page 1
    Row[4]:     data page 2
    Row[n]:     data page n

    Multiple values:

    If a propertie has multiple values such as "Has temporal scale" just make as much colomns
    as the propertie has values. The parser will automatically grab and comma-separate these values.

    Site name:

    You can specify the wikipage Name by adding a propertie called "WikiPageName". If WikiPageName doesn't exist, the parser will
    name the files/pages according to their row number. The default name is "Defaul_semanticParser_" followed by the row number.

Sample Files
============

    The sample file is needed to generate the output. The generated source code will contain each propertie specified and sorted by the
    sample file.

    A sample file is just a blank set of source code with all expected properties.

    For example:

    {{Methods and Dimensions
    |Has year=
    |Has ISI=
    |Has reviewer=
    |Has temporal scale=
    |Has spatial context=
    |Has spatial scale=
    |Has objectives dimension=
    |Has goods and services dimension=
    |Has decision making dimension=
    |Has advantages=
    |Has disadvantages=
    |Has main contraints=
    |Has related DSS development=
    |Has related DSS=
    }}

    Important:

    - The order of the properties can easily set by changing the order of the sample.

    - The sample files need to be in the Sample/ directory. 
    
    - The parser will only append a propertie in the source
      code if it is included in the sample!

    - But if you included a propertie which is not contained in the excel list,
      it wont get appended either.

The Rules
=========

    Rules are just python dictionarys containing the propertie names and their wiki-equivalent.
    So if you want to parse an excel file where the propertie names aren't the same as the semanticwiki-properties you can
    link them by adding them to the rule dictionary.

    For example:

    If a propertie is called "Responsible organisation=" but for the wiki it should be named "|Has responsible organisation=" 
    you have to generate the rule as follows:

    rule = {"Responsible organisation=" : "|Has responsible organisation="}
    
    For each propertie there must be a entry in the rule dictionary to tell the parser how the propertie is called in the excel sheet and
    in the wiki source code.

    Important:

    - It is necessary to specify a rule to generate output. If the rule is blank, the output will be filtered!


Further Informations
====================

    For further informations just read the docstrings in the source code.
