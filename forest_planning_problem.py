#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Script for pushing forest planning problem data from country reports
# to the FORSYS wiki
import sys
sys.path.append('mwclient')
import os
import mwclient
import xlrd
from copy import copy
import logging
FORMAT = '%(asctime)-15s %(message)s'
logging.basicConfig(level=logging.INFO, format=FORMAT)

# configuration parameters, check hostname!
#WIKI = 'fp0804.emu.ee'
WIKI = 'test.forsys.siwawa.org'
API_PATH = '/wiki/'
ACTIONS = ['create_pages', ]

PLANNING_PROBLEM_SHEET = 'Country_ProblemType'
RELATED_DSS_SHEET = 'CountryProblemType_DSS'
MODEL_SHEET = 'CountryProblemType_Models'
METHOD_SHEET = 'CountryProblemType_Methods'
KM_PROCESS_SHEET = 'CountryProblemType_KMProcess'
KM_TECHNIQUE_SHEET = 'CountryProblemType_KMTech'
PP_TECHNIQUE_SHEET = 'CountryProblemType_PPTech'
PP_TASK_SHEET = 'CountryProblemType_PPTasks'

MAIN_TEMPL = \
'{{Forest planning problem type\n'\
'|Has full name=%(Country)s.%(ProbType)s\n'\
'|Has temporal scale=%(TempScale)s\n'\
'|Has spatial context=%(SpatContext)s\n'\
'|Has spatial scale=%(SpatScale)s\n'\
'|Has objectives dimension=%(Objective)s\n'\
'|Has goods and services dimension=%(Goods&Services)s\n'\
'|Has decision making dimension=%(PartInvolved)s\n'\
'|Has country=%(Country)s\n'\
'|Has decision support techniques=%(Country)s.%(ProbType)s.Decision_support_techniques\n'\
'|Has knowledge management processes=%(Country)s.%(ProbType)s.Knowledge_management_processes\n'\
'|Has support for social participation=%(Country)s.%(ProbType)s.Social_participation\n'\
'|Has related DSS=%(DSCDSS)s\n'\
'}}\n'

DECISION_SUPPORT_TEMPL = \
'{{Decision support techniques\n'\
'|Has forest model=%(Simulation)s\n'\
'|Has ecological model=%(Ecological)s\n'\
'|Has social model=%(social model)s\n'\
'|Has MCDM method=%(MCDM)s\n'\
'|Has optimisation package=%(optimisation package)s\n'\
'|Has optimisation algorithm=%(Optimisation)s\n'\
'|Has risk evaluation=%(risk evaluation)s\n'\
'|Has uncertainty evaluation=%(uncertainty evaluation)s\n'\
'|Has planning scenario=%(planning scenario)s\n'\
'|Has other models=%(Other)s\n'\
'}}\n'

DECISION_SUPPORT_DEF = {
'Simulation': '',
'Ecological': '',
'social model': '',
'MCDM': '',
'optimisation package': '',
'Optimisation': '',
'risk evaluation': '',
'uncertainty evaluation': '',
'planning scenario': '',
'Other': '',
}

KM_TEMPL = \
'{{Knowledge management process\n'\
'|supports KM process=%(KM Process)s\n'\
'|Has KM techniques to identify and structure knowledge=%(identify)s\n'\
'|Has KM techniques to transfer and share knowledge=%(transfer)s\n'\
'|Has KM techniques to analyse and apply knowledge=%(analyse)s\n'\
'|Has KM techniques to unspecificed process=%(Integrated KM techniques to '\
    'unspecificed process)s\n'\
'}}\n'

KM_DEF = {
'KM Process': '',
'identify': '',
'transfer': '',
'analyse': '',
'Integrated KM techniques to unspecificed process': '',
}

SOCIAL_TEMPL = \
'{{Social participation process\n'\
'|Has participatory planning task=%(Participatory Planning Tasks)s\n'\
'|Has participatory planning techniques=%(Participatory Planning Techniques)s\n'\
'|Supports stakeholder identification=%(stakeholder identification)s\n'\
'|Supports planning criteria formation=%(criteria formation)s\n'\
'|Supports planning process monitoring and evaluation=%(process monitoring)s\n'\
'|Supports planning outcome monitoring and evaluation=%(outcome monitoring)s\n'\
'|Has stakeholder involvement=%(stakeholder involvement)s\n'\
'}}\n'

SOCIAL_DEF = {
'Participatory Planning Tasks': '',
'Participatory Planning Techniques': '',
'stakeholder identification': '',
'criteria formation': '',
'process monitoring': '',
'outcome monitoring': '',
'stakeholder involvement': '',
}

class SemanticPageBot(object):

    def __init__(self, username, pwd, url, path):
        """
        Constructor

        url -- basepath for the wiki to edit
        path -- path to the wiki API location
        """
        self.site = mwclient.Site(url, path)
        self.site.login(username, pwd)


    def create_forest_planning_problem_pages(self, excel):
        """
        Create a forest planning problem page and its sub pages

        """
        self.wb = xlrd.open_workbook(excel)
        self._get_planning_problems(PLANNING_PROBLEM_SHEET)
        rel_dss = self._get_other_data(RELATED_DSS_SHEET)
        for key in rel_dss.keys():
            self.planning_problem[key]['DSCDSS'] = rel_dss[key]['DSCDSS']
        for key in self.planning_problem.keys():
            if 'DSCDSS' not in self.planning_problem[key]:
                self.planning_problem[key]['DSCDSS'] = ''

        ds_data = {}
        self._get_multiple_values(ds_data, MODEL_SHEET, 'TypeOfModel',
                              'Description of a Model', DECISION_SUPPORT_DEF)
        self._get_multiple_values(ds_data, METHOD_SHEET, 'Type of Method',
                                  'Method', DECISION_SUPPORT_DEF)

        km_data = {}
        self._get_multiple_values(km_data, KM_PROCESS_SHEET, 'TypeOfProcess',
                                  'KM Process', KM_DEF)
        self._get_multiple_values(km_data, KM_TECHNIQUE_SHEET,
                                  'Support of Knowldge Management',
                                  'KM Technique', KM_DEF)

        sp_data = {}
        self._get_multiple_values(sp_data, PP_TECHNIQUE_SHEET,
                                  'TypeOfTechnique',
                                  'Participatory Planning Techniques',
                                  SOCIAL_DEF)
        self._get_multiple_values(sp_data, PP_TASK_SHEET,
                                  'TypeOfTask',
                                  'Participatory Planning Tasks', SOCIAL_DEF)

        for key, pp in self.planning_problem.iteritems():
            page_name = '%s.%s' % key
            self._push_page(page_name, MAIN_TEMPL,  pp)

            ds_name = page_name + '.Decision_support_techniques'
            if not key in ds_data:
                ds_data[key] = copy(DECISION_SUPPORT_DEF)
            self._push_page(ds_name, DECISION_SUPPORT_TEMPL, ds_data[key],
                            True)

            km_name = page_name + '.Knowledge_management_processes'
            if not key in km_data:
                km_data[key] = copy(KM_DEF)
            self._push_page(km_name, KM_TEMPL, km_data[key], True)

            sp_name = page_name + '.Social_participation'
            if not key in sp_data:
                sp_data[key] = copy(SOCIAL_DEF)
            self._push_page(sp_name, SOCIAL_TEMPL, sp_data[key], True)


    def _push_page(self, page_name, template, page_data, multivalue=False):
        """
        Save a page in wiki

        page_name -- page name
        template -- page content template to use
        page_data -- page data
        multivalue -- page data contains lists of values
        """
        if multivalue:
            for key, vals in page_data.iteritems():
                page_data[key] = ','.join(vals)
        page_text = template % page_data
        logging.info('Pushing page %s to wiki' % page_name)
        logging.info(page_text)
        logging.info('-' * 20)
        page_name.replace(' ', '_')
        page = self.site.pages['%s' % page_name]
        page.save(page_text, summary='')


    def _get_planning_problems(self, sheet_name):
        """
        Get planning problem data

        sheet_name -- name of the sheet containing the data
        """
        sh = self.wb.sheet_by_name(sheet_name)
        headers = sh.row_values(0)
        self.planning_problem = {}
        for rownum in range(1, sh.nrows):
            vals = sh.row_values(rownum)
            row_dict = dict(zip(headers, vals))
            if 'ProbType' in row_dict:
                row_dict['ProbType'] = int(row_dict['ProbType'])
            for key, val in row_dict.iteritems():
                try:
                    new_vals = val.split(';')
                    if len(new_vals) > 1:
                        new_val = ','.join([v.strip() for v in new_vals \
                                           if v not in ('', ' ', '  ', '   ')])
                        row_dict[key] = new_val
                except:
                    pass
            key = (row_dict['Country'], row_dict['ProbType'])
            self.planning_problem[key] = row_dict

    def _get_other_data(self, sheet_name):
        """
        Get other data related to the planning problem data

        sheet_name -- name of the sheet containing the data
        """
        sh = self.wb.sheet_by_name(sheet_name)
        headers = sh.row_values(0)
        other_data = {}
        for rownum in range(1, sh.nrows):
            vals = sh.row_values(rownum)
            row_dict = dict(zip(headers, vals))
            if 'ProbType' in row_dict:
                row_dict['ProbType'] = int(row_dict['ProbType'])
            key = (row_dict['Country'], row_dict['ProbType'])
            other_data[key] = row_dict
        return other_data

    def _get_multiple_values(self, data, sheet_name, g_key, v_key, def_dict):
        """
        Get data related to the planning problem data from
        several rows for each planning problem

        data -- collected data stored here
        sheet_name -- name of the sheet containing the data
        g_key -- column header for the grouping attribute
        v_key -- column header for the value attribute
        def_dict -- dictionary containing all the keys with empty values
        """
        sh = self.wb.sheet_by_name(sheet_name)
        headers = sh.row_values(0)
        for rownum in range(1, sh.nrows):
            vals = sh.row_values(rownum)
            row_dict = dict(zip(headers, vals))
            if 'ProbType' in row_dict:
                row_dict['ProbType'] = int(row_dict['ProbType'])
            key = (row_dict['Country'], row_dict['ProbType'])
            if key not in data:
                data[key] = copy(def_dict)
            grouper = row_dict[g_key]
            if data[key][grouper] == '':
                data[key][grouper] = []
            data[key][grouper].append(row_dict[v_key])


def main(options, username, pwd, excel, action):
    bot = SemanticPageBot(username, pwd, WIKI, API_PATH)
    if action == 'create_pages':
        bot.create_forest_planning_problem_pages(excel)

if __name__ == '__main__':
    from optparse import OptionParser
    usage = "usage: %prog [options] username:password excel_file action\n"\
            "    excel_file -- FORSYS_WIKI_properties from Google Docs\n"\
            "    action -- create_pages\n"
    parser = OptionParser(usage)
    parser.add_option('-d', '--debug', dest='debug', action='store_true',
                      help='Run post mortem debug mode enabled')
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
    if not os.path.exists(excel):
        parser.error('The file %s does not exist' % excel)
    if action not in ACTIONS:
        parser.error('Action must be one of: %s' % ACTIONS)

    main(options, username, pwd, excel, action)
