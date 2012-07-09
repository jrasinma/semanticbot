semanticbot
===========

Semantic data integration between FORSYS Google Docs documents and FORSYS semantic mediawiki

How to use
==========

1.) Change the hostname in the python script (without "http://", for example: WIKI = 'fp0804.emu.ee')

2.) Push Properties
  -> change username and password (and excel file name)
  
  command line:
  semanticbot.py username:password forsys_semanticwiki_properties.xls push_properties

3.) Create Form
  -> check if CREATE_TEMPLATES=True or False
  -> use the last parameter for cut_off, '0' means that all the properties will be used, otherwise it depends on the properties priority (see google docs)
  
  command line:
  semanticbot.py username:password forsys_semanticwiki_properties.xls create_form 0