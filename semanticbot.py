#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Script for editing the FORSYS wiki
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

WIKI = 'hostname'
API_PATH = '/wiki/'
ACTIONS = ['push_properties', 'list_priorities', 'create_form']
CREATE_TEMPLATES = False

TARGET_FORM_HEADER = u'Target form'
ID_HEADER = u'Property ID'
NAME_HEADER = u'Property'
TYPE_HEADER = u'Type'
META_HEADER = u'Metadata'
TOOLTIP_HEADER = u'Tooltip'
ENUM_HEADER = u'Value'
PRIORITY_HEADER = u'Priority score'
CRITERIA_HEADER = u'Criteria'
UI_HEADER = u'Form control'
MULTIP_HEADER = u'Form multiplicity'
MANDATORY_HEADER = u'Mandatory'
DEPENDENCY_HEADER = u'Depends on'

# In which order are the properties place on the form from the different
# 'Criteria' sections on the 'forsys_semanticwiki_properties' sheet
# given by form names from the 'Target form' column
FORM_ORDER = {'DSS': (u'Wiki quality control',
                      u'Name, responsible organisation and contact person',
                      u'Scope of the tool',
                      u'Concrete application',
                      u'Installation and support',
                      u'Data, data model and data management',
                      u'Models and methods, MBMS, decision support techniques',
                      u'Support of knowledge management process',
                      u'Support of social participation',
                      u'User interface and outputs',
                      u'System design and development',
                      u'Technological architecture, integration with other systems',
                      u'Ongoing development',
                      u'Documentation',
                     ),
             'Case': (u'Wiki quality control',
                      u'Name, responsible organisation and contact person',
                      u'Scope of the tool',
                      u'Concrete application',
                      u'Installation and support',
                      u'Data, data model and data management',
                      u'Models and methods, MBMS, decision support techniques',
                      u'Support of knowledge management process',
                      u'Support of social participation',
                      u'User interface and outputs',
                      u'System design and development',
                      u'Technological architecture, integration with other systems',
                      u'Ongoing development',
                      u'Documentation',
                      u'Action',
                      u'Case study',
                      u'Lessons-learned',
                      u'Guidelines',
                     ),
             }

TEMPLATE_TEMPL = \
'<noinclude>\n'\
'This is the "%(name)s" template.\n'\
'It should be called in the following format:\n'\
'<pre>\n'\
'{{%(name)s\n'\
'%(pre)s'\
'|Additional information=\n'\
'}}\n'\
'</pre>\n'\
'Edit the page to see the template text.\n'\
'</noinclude><includeonly>==== %(category)s ====\n'\
'{| class="wikitable dsstable"\n'\
'%(struct)s'\
'! Additional information\n'\
'| {{{Additional information|}}}\n'\
'|-\n'\
'|}\n\n'\
'[[Category:%(form_name)s]]\n'\
'</includeonly>\n'

TEMPL_PRE = '|%s=\n'
TEMPL_SINGLE = '! %(name)s\n'\
               '| [[%(name)s::{{{%(name)s|}}}]]\n'\
               '|-\n'
TEMPL_MULTIPLE = '! %(name)s\n'\
                 '| {{#arraymap:{{{%(name)s|}}}|,|@@@@|[[%(name)s::@@@@]]}}\n'\
                 '|-\n'

FORM_BEGIN = \
'<noinclude>\n'\
'This is the "%(form_name)s" form.\n'\
'To create a page with this form, enter the page name below;\n'\
'if a page with that name already exists, you will be sent to a form to edit '\
'that page.\n\n\n'\
'{{#forminput:form=%(form_name)s}}\n\n'\
'</noinclude><includeonly>\n'\
'<div id="wikiPreview" style="display: none; padding-bottom: 25px; '\
'margin-bottom: 25px; border-bottom: 1px solid #AAAAAA;"></div>\n'

FORM_TEMPL = \
'{{{for template|%(name)s|label=%(label)s}}}\n'\
'{| class="formtable"\n'\
'%(fields)s'\
'! Additional information: {{#info: Some comment with Wiki-Syntax}}\n'\
'| {{{field|Additional information|input type=textarea}}}\n'\
'|-\n'\
'%(extra_tooltip)s'\
'|}\n'\
'%(divs)s\n'\
'{{{end template}}}\n\n'

FORM_END = \
'\'\'\'Free text: {{#info: Free text with wiki-syntax. '\
'Use the following tag to list up references used in this document: '\
'"<nowiki><references/></nowiki>"}}\'\'\'\n\n'\
'{{{standard input|free text|rows=25|cols=110}}}\n\n\n'\
'{{{standard input|summary}}}\n\n'\
'{{{standard input|minor edit}}} {{{standard input|watch}}}\n'\
'{{{standard input|save}}} {{{standard input|preview}}} '\
'{{{standard input|changes}}} {{{standard input|cancel}}}\n'\
'</includeonly>\n'

FORM_SINGLE = '! %(name)s: {{#info: %(tooltip)s}}\n'\
              '| {{{field|%(name)s%(mandatory)s%(triggers)s}}}\n'\
              '|-\n'
FORM_MULTIPLE = '! %(name)s: {{#info: %(tooltip)s}}\n'\
                '| {{{field|%(name)s|input type=%(ctrl)s%(mandatory)s'\
                '%(default)s%(size)s%(maxlength)s%(triggers)s}}}\n'\
                '|-\n'
FORM_DIV_SINGLE = '! %(name)s:\n'\
              '| {{{field|%(name)s%(mandatory)s%(triggers)s}}}\n'\
              '|-\n'
FORM_DIV_MULTIPLE = '! %(name)s:\n'\
                '| {{{field|%(name)s|input type=%(ctrl)s%(mandatory)s|'\
                '%(default)s%(size)s%(maxlength)s%(triggers)s}}}\n'\
                '|-\n'
FORM_CONDITIONAL = '<div id="%(id)s">\n{|\n%(field)s|}\n</div>\n'
EXTRA_TOOLTIP = '+++%(name)s+++: \n%(tooltip)s\n\n'


class SemanticBot(object):

    def __init__(self, username, pwd, url, path):
        """
        Constructor

        url -- basepath for the wiki to edit
        path -- path to the wiki API location
        """
        self.site = mwclient.Site(url, path)
        self.site.login(username, pwd)

    def _create_property_wiki_definition(self, row_dict, data_type):
        """
        Create the semantic wiki definition for the property

        row_dict -- data from Excel property sheet
        data_type -- data type of the property
        """
        t_type = u'This is a property of type [[Has type::%s]].'
        t_meta = u'\n\n%s'
        t_enums = u'\n\nThe allowed values for this property are:'
        t_enum = u'\n* [[Allows value::%s]]'

        # property semantic wiki definition
        p_def = t_type % data_type
        meta = row_dict[META_HEADER]
        tooltip = row_dict[TOOLTIP_HEADER]
        if not tooltip:
            tooltip = meta
        p_meta = ''
        if meta != '':
            p_def += t_meta % meta
            p_meta = meta

        # process the enumeration of possible values for the property
        enum = row_dict[ENUM_HEADER]
        ctrl = row_dict[UI_HEADER].lower()
        mandatory = row_dict[MANDATORY_HEADER]
        enums = []
        if enum != '':
            p_def += t_enums
            enums = [v.strip() for v in enum.split(',')]
            if  ctrl == u'radiobutton' and not mandatory and u'N/A' not in enums:
                enums.insert(0, u'N/A')
            for e in enums:
                p_def += t_enum % e

        return p_def, p_meta, tooltip, enums

    def read_property_Excel(self, e_file):
        """
        Reads the semantic property definitions from the given Excel file
        conforming to the structure defined with the *_HEADER constants

        e_file -- Excel file to process
        """
        wb = xlrd.open_workbook(e_file)
        sh = wb.sheet_by_index(0)
        headers = sh.row_values(0)
        props = []
        trigger_p = {}
        old_criteria = ''
        ordinal = 0
        Property = namedtuple('Property',
                              'target_forms id name definition meta tooltip '\
                              'priority criteria ui_control form_order '\
                              'multiplicity mandatory default div_name')
        for rownum in range(1, sh.nrows):
            vals = sh.row_values(rownum)
            row_dict = dict(zip(headers, vals))
            data_type = row_dict[TYPE_HEADER]
            if data_type == '':
                # data type column is used to infer whether this is really a
                # property definition row; if no data type, not a property row
                continue
            res = self._create_property_wiki_definition(row_dict, data_type)
            p_def, p_meta, tooltip, enums = res
            tfs = row_dict[TARGET_FORM_HEADER]
            target_forms = [v.strip() for v in tfs.split(',')]
            p_id = int(row_dict[ID_HEADER])
            name = row_dict[NAME_HEADER]
            priority = row_dict[PRIORITY_HEADER]
            criteria = row_dict[CRITERIA_HEADER]
            if old_criteria != criteria:
                # new section of properties
                old_criteria = criteria
                ordinal = 0
            ordinal += 1
            multip = row_dict[MULTIP_HEADER]

            # fixes for handling radiobuttons
            ctrl = row_dict[UI_HEADER].lower()
            mandatory = row_dict[MANDATORY_HEADER]
            if mandatory or ctrl == u'radiobutton':
                mandatory = '|mandatory'
            default = ''
            if ctrl == u'radiobutton' and len(enums) > 0:
                default = '|default=' + enums[0]

            # process properties depending on some other property and its value
            depends_on = row_dict[DEPENDENCY_HEADER]
            div_name = ''
            if depends_on:
                # push the dependent properties to the master property
                # show on select setting
                triggers = [v.strip() for v in depends_on.split(',')]
                for t in triggers:
                    t_id, selection = [v.strip() for v in t.split(';')]
                    t_id = int(t_id)
                    if t_id not in trigger_p:
                        trigger_p[t_id] = '|show on select='
                    div_name = name.replace(' ', '_')
                    trigger_p[t_id] += '%s=>%s;' % (selection, div_name)

            props.append(Property(id=p_id,
                                  target_forms=target_forms,
                                  name=name,
                                  definition=p_def,
                                  meta=p_meta,
                                  tooltip=tooltip,
                                  priority=priority,
                                  criteria=criteria,
                                  ui_control=ctrl,
                                  form_order=ordinal,
                                  multiplicity=multip,
                                  mandatory=mandatory,
                                  default=default,
                                  div_name=div_name))
        self.properties = props
        self.trigger_properties = trigger_p


    def put_properties2wiki(self):
        """
        Define properties in the wiki
        """
        for p in self.properties:
            logging.info('Pushing property:')
            logging.info(p.name)
            logging.info(p.definition)
            logging.info('-' * 20)
            page = self.site.pages['Property:%s' % p.name]
            summary = 'Definition of property from the WG1 property sheet'
            page.save(p.definition, summary=summary)

    def list_priorities(self):
        """
        List the properties in decreasing priority score order
        """
        p_order = []
        for p in self.properties:
            p_order.append((p.priority, p.name, p.criteria))
        p_order.sort(reverse=True)
        pprint(p_order)

    def _add_non_cond_form_prop(self, p, triggers):
        """
        Add a non conditional property, endin in the form itself

        p -- property definition
        triggers -- properties to be shown when this is selected
        """
        if p.ui_control and p.ui_control != 'checkbox':
            size = ''
            maxlength = ''
            if ';' in p.ui_control:
                ctype, value = p.ui_control.split(';')
                if ctype != 'textarea':
                    size = '|size=%s' % value
                else:
                    maxlength = '|maxlength=%s' % value
            else:
                ctype = p.ui_control
            sd = {'name': p.name, 'tooltip': p.tooltip,
                  'ctrl': ctype, 'size': size, 'maxlength': maxlength,
                  'mandatory': p.mandatory, 'default': p.default,
                  'triggers': triggers}
            field_def = FORM_MULTIPLE % sd
        else:
            sd = {'name': p.name, 'tooltip': p.tooltip,
                  'mandatory': p.mandatory, 'triggers': triggers}
            field_def = FORM_SINGLE % sd
        return field_def

    def _add_cond_form_prop(self, p, triggers):
        """
        Add a conditional property ending to a separate div
        NB! Tooltips don't work property inside the divs

        p -- property definition
        triggers -- properties to be shown when this is selected
        """
        if p.ui_control and p.ui_control != 'checkbox':
            size = ''
            maxlength = ''
            if ';' in p.ui_control:
                ctype, value = p.ui_control.split(';')
                if ctype != 'textarea':
                    size = '|size=%s' % value
                else:
                    maxlength = '|maxlength=%s' % value
            else:
                ctype = p.ui_control
            sd = {'name': p.name,
                  'ctrl': ctype, 'size': size, 'maxlength': maxlength,
                  'mandatory': p.mandatory, 'default': p.default,
                  'triggers': triggers}
            field_def = FORM_DIV_MULTIPLE % sd
        else:
            sd = {'name': p.name,
                  'mandatory': p.mandatory, 'triggers': triggers}
            field_def = FORM_DIV_SINGLE % sd
        sd = {'id': p.div_name, 'field': field_def}
        return sd

    def _create_template(self, form_name, templ_name, pre, category,
                         templ_struct):
        """
        Creates a template definition and stores it in the wiki

        form_name -- form being created
        templ_name -- template name
        pre -- description text
        category -- category the template is for
        templ_struct -- template definition
        """
        sd = {'form_name': form_name, 'name': templ_name, 'pre': pre,
              'category': category, 'struct': templ_struct}
        template = TEMPLATE_TEMPL% sd
        logging.info('Pushing template %s to wiki' % templ_name)
        logging.info(template)
        logging.info('-' * 20)
        page_name = templ_name.replace(' ', '_')
        page = self.site.pages['Template:%s' % page_name]
        summary = 'Definition of template derived from the WG1 '\
                  'property sheet'
        page.save(template, summary=summary)

    def _collect_category_properties(self, form_name, category, form_p):
        """
        Process properties for the given category for inclusion in the form

        form_name -- the name of the form being created
        category -- category to get properties for
        form_p -- all properties
        """
        templ_struct = ''
        pre = ''
        templ_name = '%s, %s' % (form_name, category)
        f_struct = ''
        divs = ''
        extra_tooltip = ''
        section_placed = False
        for p in form_p:
            if form_name not in p.target_forms or p.criteria != category:
                continue
            if not section_placed:
                #f_struct += SECTION_TEMPL % category
                section_placed = True
            # Form part
            triggers = self.trigger_properties.get(p.id, '')
            if triggers:
                triggers = triggers[:-1]
            if not p.div_name:
                field_def = self._add_non_cond_form_prop(p, triggers)
                f_struct += field_def
            else:
                sd = self._add_cond_form_prop(p, triggers)
                divs += FORM_CONDITIONAL % sd
                # tooltip to be placed outside the div
                sd = {'name': p.name, 'tooltip': p.tooltip}
                extra_tooltip += EXTRA_TOOLTIP % sd

            # Template part
            pre += TEMPL_PRE % p.name
            if p.multiplicity == u'single':
                templ_struct += TEMPL_SINGLE % {'name':p.name}
            else:
                templ_struct += TEMPL_MULTIPLE % {'name':p.name}
        return (section_placed, templ_struct, pre, templ_name, f_struct, divs,
                extra_tooltip)

    def create_form(self, form_name, cut_off):
        """
        Create the DSS description form and the template it's base on out of
        the properties in the sheet based on the priority scores given for
        the properties

        form_name -- name of the form to create, must match the name on the
                     'Target form' column in the property sheet
        cut_off -- cut off for property priority score for inclusion in the
                   form
        """
        form_p = []
        for p in self.properties:
            if p.priority >= cut_off:
                form_p.append(p)
        form_p.sort(key=lambda prop: (prop.criteria, prop.form_order))
        form_templates = ''
        for category in FORM_ORDER:
            res = self._collect_category_properties(form_name, category,
                                                    form_p)

            section_placed, templ_struct, pre, templ_name, f_struct, divs,\
                extra_tooltip = res

            if section_placed:
                if CREATE_TEMPLATES:
                    self._create_template(form_name, templ_name, pre, category,
                                          templ_struct)
                # form
                if extra_tooltip:
                    extra_tooltip = '!+ {{#info: %s}}\n' % extra_tooltip
                sd = {'name': templ_name, 'label': category,
                      'fields': f_struct, 'divs': divs,
                      'extra_tooltip': extra_tooltip}
                form_templates += FORM_TEMPL % sd

        begin = FORM_BEGIN % {'form_name': form_name}
        form = begin + form_templates + FORM_END
        logging.info('Pushing form %s to wiki' % form_name)
        logging.info(form)
        logging.info('-' * 20)
        page_name = form_name.replace(' ', '_')
        page = self.site.pages['Form:%s' % page_name]
        summary = 'Definition of form derived from the WG1 property sheet'
        page.save(form, summary=summary)

def main(options, username, pwd, excel, action, form_name, cut_off):
    bot = SemanticBot(username, pwd, WIKI, API_PATH)
    bot.read_property_Excel(excel)
    if action == 'push_properties':
        bot.put_properties2wiki()
    elif action == 'list_priorities':
        bot.list_priorities()
    elif action == 'create_form':
        bot.create_form(form_name, cut_off)

if __name__ == '__main__':
    from optparse import OptionParser
    usage = "usage: %prog [options] username:password excel_file action\n"\
            "    excel_file -- FORSYS_WIKI_properties from Google Docs\n"\
            "    action -- push_properties, list_priorities, create_form\n"\
            "    cut_off -- only for create_form: priority cut off score"
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
    if action == 'create_form':
        try:
            form_name = args[3]
        except:
            parser.error('Give the form name when creating form')
        try:
            cut_off = float(args[4])
        except:
            parser.error('Give the cut off priority score when creating form')
    if not os.path.exists(excel):
        parser.error('The file %s does not exist' % excel)
    if action not in ACTIONS:
        parser.error('Action must be one of: %s' % ACTIONS)

    main(options, username, pwd, excel, action, form_name, cut_off)
