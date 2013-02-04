#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Script for importing empirical guidelines survey results as lessons to
# the FORSYS wiki
import sys
sys.path.append('mwclient')
import os
import mwclient
import xlrd
from collections import namedtuple
from pprint import pprint
import logging
FORMAT = '%(asctime)-15s %(message)s'
logging.basicConfig(level=logging.INFO, format=FORMAT)

# configuration parameters, check hostname!
WIKI = 'fp0804.emu.ee'
API_PATH = '/wiki/'
ACTIONS = ['push_lessons', ]

# The mapping of Excel column headers to Lesson form properties given here
# in [Property name=%(Excel column name)s format
LESSON_TEMPLATE = '''{{Lesson
|Domain=%(Domain)s
|Topic=%(Topic)s
|What=%(What)s
|How=%(How)s
|Why=%(Why)s
}}
'''

IS_LESSON_HEADER = u'??'

class SemanticBot(object):

    def __init__(self, username, pwd, url, path):
        """
        Constructor

        url -- basepath for the wiki to edit
        path -- path to the wiki API location
        """
        self.site = mwclient.Site(url, path)
        self.site.login(username, pwd)

    def read_lesson_Excel(self, e_file):
        """
        Reads the lesson definitions from the given Excel file
        conforming to the structure defined with the *_HEADER constants

        e_file -- Excel file to process
        """
        wb = xlrd.open_workbook(e_file)
        sh = wb.sheet_by_index(0)
        headers = sh.row_values(0)
        l_names = set([])
        lessons = []
        ordinal = 0
        Lesson = namedtuple('Lesson', 'name definition')
        name_error = False
        for rownum in range(1, sh.nrows):
            vals = sh.row_values(rownum)
            row_dict = dict(zip(headers, vals))
            is_lesson = row_dict[IS_LESSON_HEADER]
            if is_lesson == '':
                # is_lesson column is used to infer whether this is really a
                # lesson definition row; if empty, not a lesson row
                continue
            name = ??
            if name in l_names:
                msg = "Lesson '%s' used at least twice in the lesson "\
                        "sheet" % name
                logging.error(msg)
                name_error = True
            l_names.add(name)
            l_def = LESSON_TEMPLATE % row_dict

            lessons.append(Lesson(name=name,
                                  definition=l_def)
        self.lessons = lessons

        if name_error:
            msg = 'Exiting due to lesson naming errors'
            logging.error(msg)
            sys.exit(1)

    def put_lessons2wiki(self):
        """
        Create lessons pages in the wiki
        """
        for p in self.lessons:
            logging.info('Pushing lesson:')
            logging.info(p.name)
            logging.info(p.definition)
            logging.info('-' * 20)
            page = self.site.pages['Property:%s' % p.name]
            summary = 'Lesson from the Thessaloniki/Zvolen empirical '\
                    'guidelines survey'
            page.save(p.definition, summary=summary)

def main(options, username, pwd, excel, action):
    bot = SemanticBot(username, pwd, WIKI, API_PATH)
    bot.read_lesson_Excel(excel)
    if action == 'push_lessons':
        bot.put_lessons2wiki()

if __name__ == '__main__':
    from optparse import OptionParser
    usage = "usage: %prog [options] username:password excel_file action\n"\
            "    excel_file -- FORSYS_WIKI_properties from Google Docs\n"\
            "    action -- push_lessons"\
    parser = OptionParser(usage)
    parser.add_option('-d', '--debug', dest='debug', action='store_true',
                      help='Run post mortem debug mode enabled')
    #parser.add_option('-u', '--username', dest='username', action='store',
    #                 help='Username used to connect to the database (will be'\
    #                 ' asked for in the interactive console if not provided)')
    options, args = parser.parse_args()
    if len(args) < 3:
        parser.error('Need to give username:password, Excel file and action')
    usr = args[0]
    if not ':' in usr:
        parser.error('Give WIKI username and password separated by :')
    else:
        username, pwd = usr.split(':')
    excel = args[1]
    action = args[2]
    form_name = None
    cut_off = None
    if not os.path.exists(excel):
        parser.error('The file %s does not exist' % excel)
    if action not in ACTIONS:
        parser.error('Action must be one of: %s' % ACTIONS)

    main(options, username, pwd, excel, action)
