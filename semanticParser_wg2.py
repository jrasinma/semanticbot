#!/usr/bin/env python
#-*- coding: utf-8 -*-
#
#  semanticBot.py - a simple parser for DSS
#
#  Copyright 2013 Aaron Schmocker <aaron@duckpond.ch>
#
#  This program is free software; you can redistribute it and/or modify
#  it under the terms of the GNU General Public License as published by
#  the Free Software Foundation; either version 2 of the License, or
#  (at your option) any later version.
# 
#  This program is distributed in the hope that it will be useful,
#  but WITHOUT ANY WARRANTY; without even the implied warranty of
#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#  GNU General Public License for more details.
# 
#  You should have received a copy of the GNU General Public License
#  along with this program; if not, write to the Free Software
#  Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston,
#  MA 02110-1301, USA.
#
#  [used VIM Properties]
#
#  :tabstop=4
#  :set list
#  :set expandtab
#
#  Uses python2.7 because mwclient has not yet portet to 3.3!
#

import os
import sys
sys.path.append('mwclient')
import xlrd
from collections import defaultdict
from io import open
import mwclient

#Constants
WIKI = 'test.forsys.siwawa.org'
API_PATH = '/wiki/'
INDEX = 'index.php?title='
SITE_NAME_DEFAULT = 'Default_semanticParser_'
SITE_NAME_SUFFIX = '.forsys'

#here are some plain untouched sample files of the test wiki.
samp_paren = open(u'Sample/samp.parent.txt', u'r')
samp_devel = open(u'Sample/samp.devel.txt', u'r')
samp_suppo = open(u'Sample/samp.support.txt', u'r')
samp_knowl = open(u'Sample/samp.knowledge.txt', u'r')
samp_decis = open(u'Sample/samp.decisionsupport.txt', u'r')
samp_softw = open(u'Sample/samp.software.txt', u'r')
samp_main = open(u'Sample/samp.main.txt', u'r')

#parent file rule       
rule_paren = {
'|notexisting'                      :'|Has flag=',      
'|Name='                            :'|Has full name=',
'|Acronym='                         :'|Has acronym=',
'|Contact person for the Wiki='     :'|Has wiki contact person=',
'|Contact e-mail for the Wiki='     :'|Has wiki contact e-mail=',
'|Description='                     :'|Has description=',
'|Modelling dimension='             :'|Has modelling scope=',
'|Temporal scale='                  :'|Has temporal scale=',
'|Spatial context='                 :'|Has spatial context=',
'|Spatial scale='                   :'|Has spatial scale=',
'|Objectives dimension='            :'|Has objectives dimension=',
'|Goods and services dimension='    :'|Has goods and services dimension=',
'|Decision making dimension='       :'|Has decision making dimension=',
'|Forest management goal='          :'|Has forest management goal=',
'|Supported tree species='          :'|Supports tree species=',
'|Supported silvicultural regime='  :'|Supports silvicultural regime=',
'|Typical use case='                :'|Has typical use case=',
'|Country='                         :'|Has country=',
'|Number of users='                 :'|Has number of users=',
'|Number of real-life applications=':'|Has number of real-life applications=',
'|Utilisation in education: kind of utilisation (demo, use)=':u'|Has utilisation in education=',
'|Tool dissemination='              :'|Has tool dissemination=',
'|notexisting'                      :'|Has decision support techniques=',
'|notexisting'                      :'|Has knowledge management processes=',
'|notexisting'                      :'|Has support for social participation=',
'|notexisting'                      :'|Has DSS development=',
'|Website='                         :'|Has website=',
'|Online demo='                     :'|Has online demo=',
'|Manual'                           :'|Has manual=',
'|Technical documentation='         :'|Has technical documentation=',
'|References='                      :'|Has reference='
}   

#development rule
rule_devel = {  
'|Software development methodology='                        :'|Has software development methodology=',
'|Development start year='                                  :'|Has development start year=',
'|Number of development years (100% equivalent)='           :'|Has development years=',
'|Development team size='                                   :'|Has development team size=',
'|Team profiles='                                           :'|Has team profile=',
'|Number of forest specialists in the development team='    :'|Has forest specialists in the development team=',
'|Number of users participating in specification='          :'|Has users participating in specification=',
'|Adaptation effort (man years)='                           :'|Has adaptation effort=',
'|KM tools used during the development of the DSS='         :'|Has KM tools applied to DSS development='
}

#support rule
rule_suppo = {  
'|Forest models='           :'|Has forest model=',
'|Ecological models='       :'|Has ecological model=',  
'|Social models='           :'|Has social model=',
'|Optimisation package='    :'|Has optimisation package=',
'|Optimisation algorithm='  :'|Has optimisation algorithm=',
'|Risk evaluation='         :'|Has risk evaluation='    ,
'|Uncertainty evaluation='  :'|Has uncertainty evaluation=',
'|Planning scenario='       :'|Has planning scenario='
}

#knowledge rule
rule_knowl = {
'|Supported KM processes='                                      :'|Supports KM process=',
'|Integrated KM techniques to analyse and apply knowledge='     :'|Has KM techniques to analyse and apply knowledge='
}

#decision rule
rule_decis = {  
'|Stakeholder identification support='          :'|Supports stakeholder identification=',
'|Planning criteria formation support='         :'|Supports planning criteria formation=',
'|Planning process monitoring and evaluation='  :'|Supports planning process monitoring and evaluation=',
'|Planning outcome monitoring and evaluation='  :'|Supports planning outcome monitoring and evaluation='
}

#software rule
rule_softw = {  
'|Responsible organisation='        :'|Has responsible organisation=',
'|Institutional framework='         :'|Has institutional framework=',
'|Contact person for the DSS='      :'|Has DSS contact person=',
'|Contact e-mail for the DSS='      :'|Has DSS contact e-mail=',
'|Status='                          :'|Has status=',
'|Accessibility='                   :'|Has accessibility=',
'|Commercial product='              :'|Is commercial product=',
'|Deployment cost='                 :'|Has deployment cost=',   
'|Installation requirements='       :'|Has installation requirements=',
'|Computational limitations='       :'|Has computational limitations=',
'|User support organization='       :'|Has user support organization=',
'|Support team size='               :'|Has support team size=',
'|Maintenance organization='        :'|Has maintenance organization=',
'|Input data requirements='         :'|Has input data requirements=',
'|Format of the input data='        :'|Has input data format=',
'|Data validation='                 :'|Has data validation=',
'|Format of the output data='       :'|Has output data format=',
'|Internal data management='        :'|Has internal data management=',
'|Database='                        :'|Has database=',
'|notexisting'                      :'|Has GIS integration=',
'|Data mining='                     :'|Has data mining=',
'|Spatial analysis='                :'|Has spatial analysis=',
'|User access control='             :'|Has user access control=',
'|Parameterised GUI='               :'|Has parameterised GUI=',
'|Map interface='                   :'|Has map interface=',
'|GUI technology='                  :'|Has GUI technology=',
'|System type='                     :'|Has system type=',
'|Application architecture='        :'|Has application architecture=',
'|Communication architecture='      :'|Has communication architecture=',
'|Operating system='                :'|Uses operating system=',
'|Programming language='            :'|Uses programming language=',
'|Scalability='                     :'|Is scalable=',
'|notexisting'                      :'|Supports interoperability=',
'|Integration with other systems='  :'|Supports integration with other systems=',
'|Price='                           :'|Has price='
}

#rule for the  main template
rule_main = {
'WikiPageTitle='                        :'|WikiPageTitle=',
'Has method='                           :'|Has method=',
'Has submethods='                       :'|Has submethods=',
'Has detailed description of methods application in the DSS=':'|Has detailed description of methods application in the DSS=',
'Has reference='                        :'|Has reference=',
'Has risk/uncertainty analysis='        :'|Has risk/uncertainty analysis=',
'Has year='                             :'|Has year=',
'Has ISI='                              :'|Has ISI=',
'Has reviewer='                         :'|Has reviewer=',
'Has main method='                      :'|Has main method=',
'Has sub method='                       :'|Has sub method=',
'Has temporal scale='                   :'|Has temporal scale=',
'Has spatial context='                  :'|Has spatial context=',
'Has spatial scale='                    :'|Has spatial scale=',
'Has decision making dimension='        :'|Has decision making dimension=',
'Has objectives dimension='             :'|Has objectives dimension=',
'Has goods and services dimension='     :'|Has goods and services dimension=',
'Has advantages='                       :'|Has advantages=',
'Has description='                      :'|Has description=',
'Has disadvantages='                    :'|Has disadvantages=',
'Has main contraints='                  :'|Has main contraints=',
'Has related DSS development='          :'|Has related DSS development=',
'Has related DSS='                      :'|Has related DSS=',
'Has related software='                 :'|Has related software=',
'Has related method='                   :'|Has related method=',
'Platform used for the development of the DSS=':'|Platform used for the development of the DSS='
}

"""
* Class Overview
*****************************************************************
* Class:        SemanticBot                                     *
* Functions:    fromFile()         : creates a file             *
*               flushSite()        : uploads content            *
*               parseExcel()       : reads an Excel             *
*               _genForsysSource() : generates content          *
*****************************************************************
* Brief:    This class contains methodes to build a semantic    *
*           wiki source code and either flush it to a file      * 
*           or directly upload it on the wiki.                  *   
*****************************************************************
"""
class SemanticBot(object):
    
    #Constructor
    def __init__(self, url, path):
        self.url = url
        self.path = path

    #login
    def mediaWikiLogin(self, username, pwd):
        self.username = username
        self.pwd = pwd
        #login
        self.site = mwclient.Site(self.url, self.path)
        self.site.login(username, pwd)
        
    """
    * Mehtod Overview
    *****************************************************************
    * Method:       fromFile                                        *
    * Parameters:   output_file_name                                *
    *               path_to_dss                                     *
    *               rule                                            *
    *               sample                                          *
    *****************************************************************
    * Brief:    This class method creates a new file and flushs     *
    *           the output of the _genForysSource method to the     *
    *           given output path.                                  *
    *****************************************************************
    """
    #flushs the output of the genForsysSource method to the given output path.
    def fromFile(self, output_file_name, path_to_dss, rule, sample):
        #generate the output file.
        output_file = open(output_file_name, u'w+')
        #this file contains all informations in a merged form.
        dss = open(path_to_dss)
        #generate string out of the dss
        dss_string = [line.strip() for line in dss]
        #write out the data.
        output_file.write(self._genForsysSource(dss_string, rule, sample))

    """
    * Method Overview
    *****************************************************************
    * Method:       uploadToSite                                    *
    * Parameters:   page_text_file_object   : File object           *
    *               page_name               : Name of the page      *
    *****************************************************************
    * Brief:    This method uploads the converted content to the    *
    *           wiki.                                               *   
    *****************************************************************
    """
    #flushs the output of the semantisize method to the wiki Site and creates a new page.
    def uploadToSite(self, page_text_file_object, page_name):
        page_name.replace(' ', '_')
        page = self.site.pages['%s' % page_name]
        page.save(page_text_file_object.read(), summary='')


    """
    * Method Overview
    *****************************************************************
    * Method:       _genForsysSource()                              *
    * Parameters:   string_to_dss   : dss as string                 *
    *               rule            : rule dictionary               *
    *               sample          : sample object                 *
    * Returns:      raw_source      : a string with the raw source  *
    *****************************************************************
    * Brief:   This method builds the raw site source out of o list *
    *          of properties.                                       *
    *          The source layout will be generated according to the *
    *          sample file and the propertie names get rematched by *
    *          the defined rule.                                    *
    *****************************************************************
    """ 
    def _genForsysSource(self, old_dss_string, rule, sample):
        #Read the dss as it should be.
        samp = [line.strip() for line in sample]
        #Read the matchin rules.
        new = [line.strip() for line in rule]
        #buffer
        buf = []
        #Initialize our output string.
        out = []
        
        #for each item in our old dss.
        for item in old_dss_string:
        
            #separate the user defined answer from the old propertie.
            pos = item.find(u'=')
            propertie = item[0:pos+1]
            userinput = item[pos+1:len(item)]
        
            #read read from the rule file, if there is no match, match contains -1.
            match = rule.get(propertie, -1)
        
            #look for the matching string.
            if match is not -1:
                buf.append(match + userinput + u'\n')
        
        #now iterate over the sample file to correct the Layout
        for item in samp:

            #if there is something with braces, append it anyway.
            if item.find('}') is not -1 or  item.find('{') is not -1:
                out.append(item + '\n')

            #append entry only to the output file if the item is contained in both buffer and sample.
            for entry in buf:
                pos = entry.find('=')
                propertie = entry[0:pos+1]
                if unicode(item) == unicode(propertie):
                    out.append(entry)
                        
        #thanks, bye
        return ''.join(out)

    """
    * Method Overview
    *****************************************************************
    * Method:       parseExcel()                                    *
    * Parameters:   excel       : path to excel file (.xls!)        *
    * Returns:      list        : containing the sourcefile paths   *
    *****************************************************************
    * Brief:    This method will fetch and read an excel sheet      *
    *           in order to extract the essential data.             *
    *           The files with the generated source code will be    *
    *           stored in the working directory and the return      *
    *           value is a list with the paths of the generated     *
    *           outputfiles.                                        *
    *****************************************************************
    """
    #parses an existing excel file and converts the informations in FORSYS_DSS form. 
    def parseExcel(self, excel):
        #constants
        PROPERTIE_ROW = 1
        DATA_ROW = 2
        #excel element and sheet
        book = xlrd.open_workbook(excel)
        sheet = book.sheet_by_index(0)
        #list for storing the properties and filenames
        proplist = []
        d = defaultdict(list)
        new = []
        createdFileList = []
        WikiPageTitle = ''

        #for each row in the excel file.
        for i, rows in enumerate(xrange(DATA_ROW, sheet.nrows)):
            #generate a list with all properties.
            for cols in xrange(sheet.ncols):
                #if the propertie is not emty 
                if unicode(sheet.cell(PROPERTIE_ROW,cols).value) is not '':
                    cell_value = sheet.cell(rows,cols).value
                    #cause excel generally handels all numbers as floats and we don't want any yearnumber
                    #represented as 2005.0, let's check if the value is a string.
                    if not isinstance(cell_value, basestring):
                        #if not explicitly convert it to an integer.
                        cell_value = int(cell_value)
                    #add value to the proplist.
                    proplist.append(sheet.cell(PROPERTIE_ROW,cols).value + '=' + unicode(cell_value))
            
            #find and merge multiple values in just one entry.
            for item in proplist:
                k,v = item.split(u'=',1)
                d[k].append(v)
            
            new = [u'{0}={1}'.format(k,', '.join(v)) for k,v in d.items()]

            #get Filename, if there is one.
            for elem in new:
                if elem.find('WikiPageTitle') is not -1:
                    prop,WikiPageTitle = elem.split(u'=',1)
            if WikiPageTitle is not '':
                filename = WikiPageTitle
            else:
                filename = SITE_NAME_DEFAULT + unicode(i) + SITE_NAME_SUFFIX

            output_file = open(filename, 'w+')
            createdFileList.append(filename)
            #reload samplefile, beacause it gets consumed by _genForsysSource.
            samp_main = open('Sample/samp.main.txt', 'r')
            #correct the data and match it with the sample file.
            output_file.write(self._genForsysSource(new, rule_main, samp_main))
            
            #clear dict and list.
            d.clear()
            proplist[:] = []
            new[:] = []
        
        #return file list
        return createdFileList

#Main method
def main(action, username, pwd, excel_path):
    
    #paths of the temporary files 
    parsedFiles = []

    #initialize bot
    bot = SemanticBot(WIKI, API_PATH)

    if action is 'upload':
        #login
        bot.mediaWikiLogin(username, pwd)
        #read Excel
        parsedFiles = bot.parseExcel(excel_path)
        print '---------------------------------------------------'
        print '|                Uploaded Files:                  |'
        print '---------------------------------------------------'
        #upload Files
        for i,item in enumerate(parsedFiles):
            #read file
            myfile = open(parsedFiles[i],'r')
            #upload it
            bot.uploadToSite(myfile, parsedFiles[i])
            #print Sitename
            print '\n' + 'Name: ' +  parsedFiles[i] + '        Size: ' + unicode(os.path.getsize(parsedFiles[i])) + ' Bytes'
            print '\n -> Link: ' + 'http://test.forsys.siwawa.org/wiki/index.php?title=' + parsedFiles[i]
    if action is 'files':
        #read Excel
        parsedFiles = bot.parseExcel(excel_path)
        print '----------------------------------------------------'
        print '|                Generated Files:                  |'
        print '----------------------------------------------------'
        #upload Files
        for i,item in enumerate(parsedFiles):
            #print summary
            print '\n' + 'Name: ' +  parsedFiles[i] + '        Size: ' + unicode(os.path.getsize(parsedFiles[i])) + ' Bytes'
#If not loaded as module 
if __name__ == '__main__':
    from optparse import OptionParser
    
    usage = "\n\n"\
            "%prog [optinos] username:password excel_file\n"\
            "excel_file  -> FORSYS_WIKI_properties in xls format\n"\
            "userdata    -> FORSYS_WIKI username:password\n\n"\
            "Options:\n\n"\
            "-u --upload\n"\
            "       Uploads content directly to the wiki by creating new pages.\n\n"\
            "-f --files\n"\
            "       Generates outputfiles and stores them in the working directory.\n"
                                    

    parser = OptionParser(usage)
    parser.add_option('-f', '--files', action='store_true')
    parser.add_option('-u', '--upload', action='store_true')
    parser.add_option('-d', '--debug',  dest='debug', action='store_true', help='Run post mortem debug mode enabled')
    
    options, args = parser.parse_args()

    if options.upload is True: 
        
        action = 'upload'

        #if not entough args passed.
        if len(args) < 2:
            parser.error('Need to give username:password and an excel file!')
        
        usr = args[0]

        #detect a lack of username or password.
        if not ':' in usr:
            parser.error('Specify WIKI username and password separated by :')
                    
        else:
            username, pwd = usr.split(':')
        
        excel = args[1]

    if options.files is True:

        action = 'files'
        username = 'dummy'
        pwd = 'pwdummy'

        if len(args) < 1:
            parser.error('No excel file specified!')
    
        excel = args[0]

    if not os.path.exists(excel):
        parser.error('The file %s does not exist' % excel)
    
    #if everything is ok, call main.
    main(action, username, pwd, excel)
