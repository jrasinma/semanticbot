#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Script for editing the FORSYS wiki
import sys
sys.path.append('mwclient')
import os
import mwclient
import xlrd
from collections import namedtuple, defaultdict
from pprint import pprint
import logging
FORMAT = '%(asctime)-15s %(message)s'
logging.basicConfig(level=logging.INFO, format=FORMAT)

# configuration parameters, check hostname!
#WIKI = 'fp0804.emu.ee'
WIKI = 'test.forsys.siwawa.org'
API_PATH = '/wiki/'
ACTIONS = ['push_properties', 'clean_properties', 'list_priorities',
           'create_form', 'create_categories']
CREATE_TEMPLATES = True

TARGET_FORM_HEADER = u'Target form'
ID_HEADER = u'Property ID'
NAME_HEADER = u'Property'
LABEL_HEADER = u'Form label'
TYPE_HEADER = u'Type'
META_HEADER = u'Metadata'
TOOLTIP_HEADER = u'Tooltip'
ENUM_HEADER = u'Value'
PRIORITY_HEADER = u'Priority score'
SECTION_HEADER = u'Title'
UI_HEADER = u'Form control'
MULTIP_HEADER = u'Form multiplicity'
MANDATORY_HEADER = u'Mandatory'
DEPENDENCY_HEADER = u'Depends on'

PAGE_PROPERTY_CATEGORY_HEADER = u'Related category for property'
PAGE_PROPERTY_CATEGORY_TYPE_HEADER = u'Type of related category'
PAGE_SPECIFIC_SUB_CATEGORY = u'specific'
GENERAL_SUB_CATEGORY = u'general'
LABEL_ADDITION4PAGE_LINK = " (name of the detail wiki page)"

# In which order are the properties place on the form from the different
# 'Title' sections on the 'forsys_semanticwiki_properties' sheet
# given by form names from the 'Target form' column
# Note that only one of the forms can be of type 'multi-template', i.e.
# consisting of different sections each having its own template

# the other forms are not multi-template, and the order in which the properties
# are placed on the form is defined with ordering in a sheet column, see
# 'form order' setting in the FORM_DEF

# 2012-07-6: u'Wiki quality control' removed at least for now since it's
# plain confusing at the moment for users


FORM_DEF = \
    {u'DSS': {'multi-template': True,
             'content':
                (u'Wiki quality control',
                 u'Name, responsible organisation and contact person',
                 u'Software identification',
                 u'Description',
                 u'Concrete application',
                 u'Decision support techniques used in the DSS',
                 u'Support of Knowledge Management',
                 u'Support of social participation',
                 u'DSS development',
                 u'DSS',
                 u'Documentation',
                ),
              'category with form': True,
           },
    u'Case study': {'multi-template': False,
                   'form order': 'FormCase',
                   'category with form': True,
                 },
    u'Lesson': {'multi-template': False,
               'form order': 'FormLesson',
               'category with form': True,
              },
     }

TEMPLATE_TEMPL = \
'<noinclude>\n'\
'This is the "%(name)s" template.\n'\
'It should be called in the following format:\n'\
'<pre>\n'\
'{{%(name)s\n'\
'%(pre)s'\
'}}\n'\
'</pre>\n'\
'Edit the page to see the template text.\n'\
'</noinclude><includeonly>==== %(category)s ====\n'\
'{| class="wikitable dsstable"\n'\
'%(struct)s'\
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

FORM_SINGLE = '! %(label)s: {{#info: %(tooltip)s}}\n'\
              '| {{{field|%(name)s%(mandatory)s%(values_from)s%(default)s'\
              '%(triggers)s'\
              '}}}\n|-\n'
FORM_MULTIPLE = '! %(label)s: {{#info: %(tooltip)s}}\n'\
                '| {{{field|%(name)s|input type=%(ctrl)s%(mandatory)s'\
                '%(values_from)s%(default)s%(size)s%(maxlength)s%(triggers)s'\
                '}}}\n|-\n'
FORM_DIV_SINGLE = '! %(label)s:\n'\
              '| {{{field|%(name)s%(mandatory)s%(values_from)s'\
              '%(default)s%(size)s%(triggers)s'\
              '}}}\n|-\n'
FORM_DIV_MULTIPLE = '! %(label)s:\n'\
                '| {{{field|%(name)s|input type=%(ctrl)s%(mandatory)s'\
                '%(values_from)s'\
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
        t_df = u'\n\n[[Has default form::%s]]'

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

        # the form that is used to edit the page this property value points
        # to
        form_category = False
        df = ''
        page_prop_category = \
                row_dict[PAGE_PROPERTY_CATEGORY_HEADER]
        page_prop_category_type = \
                row_dict[PAGE_PROPERTY_CATEGORY_TYPE_HEADER].lower()
        if page_prop_category_type == PAGE_SPECIFIC_SUB_CATEGORY:
            df = page_prop_category
            has_def_form = '[[[Has default form::%s]]' % page_prop_category
        elif page_prop_category_type == GENERAL_SUB_CATEGORY:
            has_def_form = ''
        if page_prop_category != '':
            self.categories[page_prop_category] = {'has def form': has_def_form,
                                                   'category pages': [],}

        # process the enumeration of possible values for the property
        enum = row_dict[ENUM_HEADER]
        ctrl = row_dict[UI_HEADER].lower()
        mandatory = row_dict[MANDATORY_HEADER]
        enums = []
        values_from = ''
        if page_prop_category_type == PAGE_SPECIFIC_SUB_CATEGORY:
            values_from = '|values from category=%s' % (page_prop_category)
        elif page_prop_category_type == GENERAL_SUB_CATEGORY:
            values_from = '|values from category=%s' % (page_prop_category)
            try:
                self.categories[page_prop_category]['category pages'] = \
                    [v.strip() for v in enum.split(';')]
            except KeyError:
                msg = "Property '%s': '%s' column has value '%s', but no "\
                        "category given in column '%s'"\
                        % (row_dict[NAME_HEADER],
                           PAGE_PROPERTY_CATEGORY_TYPE_HEADER,
                           GENERAL_SUB_CATEGORY,
                           PAGE_PROPERTY_CATEGORY_HEADER)
                print msg
                sys.exit(1)
        elif enum.startswith('property:'):
            values_from = '|values from property=%s' % (enum.split(':')[1],)
        elif enum.startswith('category:'):
            values_from = '|values from category=%s' % (enum.split(':')[1],)
        elif enum != '':
            p_def += t_enums
            enums = [v.strip() for v in enum.split(';')]
            if  ctrl == u'radiobutton' and not (mandatory and u'N/A' in enums):
                enums.insert(0, u'N/A')
            for e in enums:
                p_def += t_enum % e

        if df != '':
            p_def += t_df % df

        return p_def, p_meta, tooltip, enums, values_from, form_category

    def read_property_Excel(self, e_file):
        """
        Reads the semantic property definitions from the given Excel file
        conforming to the structure defined with the *_HEADER constants

        e_file -- Excel file to process
        """
        self.categories = {}
        self.form_def = FORM_DEF
        wb = xlrd.open_workbook(e_file)
        sh = wb.sheet_by_index(0)
        headers = sh.row_values(0)
        p_names = set([])
        props = []
        trigger_p = {}
        old_criteria = ''
        ordinal = 0
        Property = namedtuple('Property',
                              'target_forms id label name definition meta '\
                              'tooltip order priority section ui_control '\
                              'dss_form_order multiplicity mandatory '\
                              'default div_name my_triggers values_from')
        name_error = False
        for rownum in range(1, sh.nrows):
            vals = sh.row_values(rownum)
            row_dict = dict(zip(headers, vals))
            data_type = row_dict[TYPE_HEADER]
            if data_type == '':
                # data type column is used to infer whether this is really a
                # property definition row; if no data type, not a property row
                continue
            res = self._create_property_wiki_definition(row_dict, data_type)
            p_def, p_meta, tooltip, enums, values_from, form_category = res
            tfs = row_dict[TARGET_FORM_HEADER]
            target_forms = [v.strip() for v in tfs.split(';')]
            for tf in target_forms:
                if tf != '' and tf not in self.form_def:
                    self.form_def[tf] = {'multi-template': False,
                                         'form order': 'sheet',
                                         'category with form': False,}
            p_id = int(row_dict[ID_HEADER])

            name = row_dict[NAME_HEADER]
            if name in p_names:
                msg = "Property '%s' used at least twice in the definition "\
                        "sheet" % name
                logging.error(msg)
                name_error = True
            p_names.add(name)
            label = row_dict[LABEL_HEADER]

            page_prop_category_type = \
                row_dict[PAGE_PROPERTY_CATEGORY_TYPE_HEADER].lower()

            if page_prop_category_type == PAGE_SPECIFIC_SUB_CATEGORY:
                label += LABEL_ADDITION4PAGE_LINK

            priority = row_dict[PRIORITY_HEADER]
            section = row_dict[SECTION_HEADER]
            order = {}
            for f_name, fdef in self.form_def.iteritems():
                if f_name == 'DSS':
                    continue
                how2order = fdef['form order']
                if how2order == 'sheet':
                    f_ord = rownum
                else:
                    f_ord = row_dict[fdef['form order']]
                if f_ord:
                    order[f_name] = f_ord
            if section == u'' and f_name == 'DSS':
                msg = 'No category/section for property %s' % name
                logging.error(msg)
                sys.exit(1)
            if old_criteria != section:
                # new section of properties
                old_criteria = section
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
            elif u'default page' in ctrl:
                default = '|default={{PAGENAME}}.%s'
                ctrl = ctrl.strip(u'default page')

            # process properties depending on some other property and its value
            depends_on = row_dict[DEPENDENCY_HEADER]
            div_name = ''
            my_triggers = set([])
            if depends_on:
                # push the dependent properties to the master property
                # show on select setting
                triggers = [v.strip() for v in depends_on.split(',')]
                for t in triggers:
                    t_id, selection = [v.strip() for v in t.split(';')]
                    t_id = int(t_id)
                    my_triggers.add(t_id)
                    if t_id not in trigger_p:
                        trigger_p[t_id] = '|show on select='
                    div_name = name.replace(' ', '_')
                    trigger_p[t_id] += '%s=>%s;' % (selection, div_name)

            props.append(Property(id=p_id,
                                  target_forms=target_forms,
                                  name=name,
                                  label=label,
                                  definition=p_def,
                                  meta=p_meta,
                                  tooltip=tooltip,
                                  order=order,
                                  priority=priority,
                                  section=section,
                                  ui_control=ctrl,
                                  dss_form_order=ordinal,
                                  multiplicity=multip,
                                  mandatory=mandatory,
                                  default=default,
                                  div_name=div_name,
                                  my_triggers=my_triggers,
                                  values_from=values_from))
        self.properties = props
        self.trigger_properties = trigger_p

        if name_error:
            msg = 'Exiting due to property naming errors'
            logging.error(msg)
            sys.exit(1)

    def clean_properties(self):
        """
        Remove properties from Wiki that are not in the property sheet
        Note, to be used with extreme caution...!
        """
        try:
            from bs4 import BeautifulSoup
            import urllib2
        except:
            msg = 'This action needs installation of BeautifulSoup module to '\
                    'function, exiting'
            logging.error(msg)
            sys.exit(1)

        wiki_props = set([])

        url = 'http://%s%sindex.php?title=Special:Properties&limit=500'\
                % (WIKI, API_PATH)
        response = urllib2.urlopen(url)
        html = response.read()
        soup = BeautifulSoup(html)
        links = soup.find_all('a')
        for l in links:
            if 'title' in l.attrs:
                if 'Property' in l.attrs['title']:
                    p = l.parent
                    if 'class' in p.attrs:
                        if u'smwbuiltin' in p.attrs[u'class']:
                            # this is a built in property
                            continue
                    #print l.attrs['title']
                    p_name = l.attrs['title'].split(':')[1]
                    if 'page does not exist' not in p_name:
                        wiki_props.add(p_name)

        sheet_props = set([])
        for p in self.properties:
            sheet_props.add(p.name)
        not_in_sheet = wiki_props.difference(sheet_props)
        if len(not_in_sheet) > 0:
            logging.info('Deleting properties: %s' % not_in_sheet)
            reason = 'Property no longer on FORSYS property sheet'
            try:
                for p in not_in_sheet:
                    page = self.site.pages['Property:%s' % p]
                    page.delete(reason)
            except:
                msg = 'Deletion failed: %s' % (p, )
                logging.error(msg)


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
            p_order.append((p.priority, p.name, p.section))
        p_order.sort(reverse=True)
        pprint(p_order)

    def _add_non_cond_form_prop(self, p, triggers, form_name):
        """
        Add a non conditional property, ending in the form itself

        p -- property definition
        triggers -- properties to be shown when this is selected
        form_name -- form name to be used in the default page name fields
        """
        if 'PAGENAME' in p.default:
            # cases that link to a default new page
            if LABEL_ADDITION4PAGE_LINK in p.label:
                last_part = p.label.replace(LABEL_ADDITION4PAGE_LINK, '')
            else:
                last_part = p.label
            default = p.default % (last_part, )
        else:
            default = p.default
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
            sd = {'label': p.label, 'name': p.name, 'tooltip': p.tooltip,
                  'ctrl': ctype, 'size': size, 'maxlength': maxlength,
                  'mandatory': p.mandatory, 'values_from': p.values_from,
                  'default': default, 'triggers': triggers}
            field_def = FORM_MULTIPLE % sd
        else:
            sd = {'label': p.label, 'name': p.name, 'tooltip': p.tooltip,
                  'mandatory': p.mandatory, 'values_from':p.values_from,
                  'default': default, 'triggers': triggers}
            field_def = FORM_SINGLE % sd
        return field_def

    def _add_cond_form_prop(self, p, triggers, form_name):
        """
        Add a conditional property ending to a separate div
        NB! Tooltips don't work property inside the divs

        p -- property definition
        triggers -- properties to be shown when this is selected
        form_name -- form name to be used in the default page name fields
        """
        if 'PAGENAME' in p.default:
            # cases that link to a default new page
            default = p.default % (p.label)
        else:
            default = p.default
        size = ''
        if p.ui_control and p.ui_control != 'checkbox':
            maxlength = ''
            if ';' in p.ui_control:
                ctype, value = p.ui_control.split(';')
                if ctype != 'textarea':
                    size = '|size=%s' % value
                else:
                    maxlength = '|maxlength=%s' % value
            else:
                ctype = p.ui_control
            sd = {'label': p.label, 'name': p.name,
                  'ctrl': ctype, 'size': size, 'maxlength': maxlength,
                  'mandatory': p.mandatory, 'values_from': p.values_from,
                  'default': p.default, 'triggers': triggers}
            field_def = FORM_DIV_MULTIPLE % sd
        else:
            sd = {'label': p.label, 'name': p.name, 'size': size,
                  'mandatory': p.mandatory, 'values_from':p.values_from,
                  'default': default, 'triggers': triggers}
            field_def = FORM_DIV_SINGLE % sd
        sd = {'id': p.div_name, 'field': field_def}
        return sd

    def _create_template(self, form_name, category, f_def):
        """
        Creates a template definition and stores it in the wiki

        form_name -- form being created
        category -- category the template is for
        f_def -- form definition data for the category
        """
        sd = {'form_name': form_name, 'name': f_def['templ_name'],
              'pre': f_def['pre'], 'category': category,
              'struct': f_def['templ_struct']}
        template = TEMPLATE_TEMPL % sd
        logging.info('Pushing template %s to wiki' % f_def['templ_name'])
        logging.info(template)
        logging.info('-' * 20)
        page_name = f_def['templ_name'].replace(' ', '_')
        page = self.site.pages['Template:%s' % page_name]
        summary = 'Definition of template derived from the WG1 '\
                  'property sheet'
        page.save(template, summary=summary)

    def _collect_all_properties(self, form_name, f_def):
        """
        Process properties for the the whole form in given order
        for inclusion in the form

        form_name -- the name of the form being created
        f_def -- form definition data
        """
        f_def['section_placed'] = True
        form_p = []
        for p in f_def['form_p']:
            if form_name not in p.target_forms:
                continue
            if form_name not in p.order:
                msg = 'Property %s labeled for form %s, but has no order '\
                        'ordinal' % (p.name, form_name)
                logging.error(msg)
                continue
            p_ordinal = p.order[form_name]
            form_p.append((p_ordinal, {'f_struct': '',
                                       'divs': '',
                                       'extra_tooltip': '',
                                       'pre': '',
                                       'templ_struct': ''}))
            # Form part
            triggers = self.trigger_properties.get(p.id, '')
            if triggers:
                triggers = triggers[:-1]
            my_trigger_exists = self._check4parent_prop(p, f_def['form_p'])
            if my_trigger_exists:
                sd = self._add_cond_form_prop(p, triggers, form_name)
                form_p[-1][1]['divs'] = FORM_CONDITIONAL % sd
                # tooltip to be placed outside the div
                sd = {'name': p.name, 'tooltip': p.tooltip}
                form_p[-1][1]['extra_tooltip'] = EXTRA_TOOLTIP % sd
            else:
                field_def = self._add_non_cond_form_prop(p, triggers,
                                                         form_name)
                form_p[-1][1]['f_struct'] = field_def

            # Template part
            form_p[-1][1]['pre'] = TEMPL_PRE % p.name
            if p.multiplicity == u'single':
                form_p[-1][1]['templ_struct'] = TEMPL_SINGLE % {'name':p.name}
            else:
                form_p[-1][1]['templ_struct'] = \
                        TEMPL_MULTIPLE % {'name':p.name}

        form_p.sort()
        for p in form_p:
            f_def['f_struct'] += p[1]['f_struct']
            f_def['divs'] += p[1]['divs']
            f_def['extra_tooltip'] += p[1]['extra_tooltip']
            f_def['pre'] += p[1]['pre']
            f_def['templ_struct'] += p[1]['templ_struct']

    def _check4parent_prop(self, prop, all_prop):
        """
        Checks whether the triggering property exists on the form for a
        property that is potentially initially hidden
        """
        trigger_exists = False
        for p in all_prop:
            if p.id in prop.my_triggers:
                trigger_exists = True
                break
        return trigger_exists

    def _collect_category_properties(self, form_name, f_def, category=None):
        """
        Process properties for the given category ('Title' on sheet)
        for inclusion in the form

        form_name -- the name of the form being created
        f_def -- form definition data
        category -- category/section to get properties for
        """
        if self.form_def[form_name]['multi-template']:
            f_def['templ_struct'] = ''
            f_def['pre'] = ''
            f_def['templ_name'] = '%s, %s' % (form_name, category)
            f_def['f_struct'] = ''
            f_def['divs'] = ''
            f_def['extra_tooltip'] = ''
            f_def['section_placed'] = False
        for p in f_def['form_p']:
            if form_name not in p.target_forms or p.section != category:
                continue
            if not f_def['section_placed']:
                f_def['section_placed'] = True
            # Form part
            triggers = self.trigger_properties.get(p.id, '')
            if triggers:
                triggers = triggers[:-1]
            my_trigger_exists = self._check4parent_prop(p, f_def['form_p'])
            if my_trigger_exists:
                sd = self._add_cond_form_prop(p, triggers, category)
                f_def['divs'] += FORM_CONDITIONAL % sd
                # tooltip to be placed outside the div
                sd = {'name': p.name, 'tooltip': p.tooltip}
                f_def['extra_tooltip'] += EXTRA_TOOLTIP % sd
            else:
                field_def = self._add_non_cond_form_prop(p, triggers, category)
                f_def['f_struct'] += field_def

            # Template part
            f_def['pre'] += TEMPL_PRE % p.name
            if p.multiplicity == u'single':
                f_def['templ_struct'] += TEMPL_SINGLE % {'name':p.name}
            else:
                f_def['templ_struct'] += TEMPL_MULTIPLE % {'name':p.name}

    def _put2form_template(self, form_name, form_templates, category, f_def):
        """
        Process info from section properties and place it to the form
        template
        """
        if f_def['section_placed']:
            if CREATE_TEMPLATES:
                self._create_template(form_name, category, f_def)
            # form
            if f_def['extra_tooltip']:
                f_def['extra_tooltip'] = '!+ {{#info: %s}}\n'\
                                         % f_def['extra_tooltip']
            sd = {'name': f_def['templ_name'], 'label': category,
                  'fields': f_def['f_struct'], 'divs': f_def['divs'],
                  'extra_tooltip': f_def['extra_tooltip']}
            form_templates += FORM_TEMPL % sd
        return form_templates

    def create_form(self, form_name, cut_off):
        """
        Create the given form and the template it's based on out of
        the properties in the sheet based on the priority scores given for
        the properties

        form_name -- name of the form to create, must match the name on the
                     'Target form' column in the property sheet
        cut_off -- cut off for property priority score for inclusion in the
                   form
        """
        if not form_name in self.form_def:
            msg = 'Form %s not in property sheet' % (form_name,)
            logging.error(msg)
            sys.exit(1)
        f_def = {'form_p': [],
                 'section_placed': False,
                 'templ_struct': '',
                 'pre': '',
                 'templ_name': '',
                 'f_struct': '',
                 'divs': '',
                 'extra_tooltip': '',
                }
        for p in self.properties:
            if cut_off is None or p.priority >= cut_off:
                f_def['form_p'].append(p)
        f_def['form_p'].sort(key=lambda prop: (prop.section,
                                               prop.dss_form_order))
        form_templates = ''
        if self.form_def[form_name]['multi-template']:
            for category in self.form_def[form_name]['content']:
                self._collect_category_properties(form_name, f_def, category)
                # form has different sections, each having its own template
                form_templates = self._put2form_template(form_name,
                                               form_templates, category, f_def)
        else:
            # the whole form in one template, the order of fields defined by
            # a column on the sheet
            self._collect_all_properties(form_name, f_def)
            f_def['templ_name'] = form_name
            category = form_name
            form_templates = self._put2form_template(form_name, form_templates,
                                                     category, f_def)

        begin = FORM_BEGIN % {'form_name': form_name}
        form = begin + form_templates + FORM_END
        logging.info('Pushing form %s to wiki' % form_name)
        logging.info(form)
        logging.info('-' * 20)
        page_name = form_name.replace(' ', '_')
        page = self.site.pages['Form:%s' % page_name]
        summary = 'Definition of form derived from the FORSYS property sheet'
        page.save(form, summary=summary)

    def create_category_pages(self):
        """
        Create categories and pages belonging to them based on the propertys
        sheet definitions
        """
        c_pages = defaultdict(str)
        for cat, cdata in self.categories.iteritems():
            #df = cdata['has def form']
            #c_text = ''
            #if df != '':
            #    c_text = 'This category uses the form %s' % (df, )
            #logging.info('Pushing category:')
            #logging.info(cat)
            #logging.info(c_text)
            #logging.info('-' * 20)
            #page = self.site.pages['Category:%s' % cat]
            #summary = 'Definition of category from the WG1 property sheet'
            #page.save(c_text, summary=summary)

            for c_page in cdata['category pages']:
                c_text = '[[Category:%s]]\n' % (cat, )
                c_pages[c_page] += c_text
        for c_page, c_text in c_pages.iteritems():
            page = self.site.pages['%s' % c_page]
            logging.info('Pushing category page:')
            logging.info(c_page)
            logging.info(c_text)
            logging.info('-' * 20)
            summary = ''
            page.save(c_text, summary=summary)


def main(options, username, pwd, excel, action, form_name):
    bot = SemanticBot(username, pwd, WIKI, API_PATH)
    bot.read_property_Excel(excel)
    if action == 'push_properties':
        bot.put_properties2wiki()
    elif action == 'clean_properties':
        bot.clean_properties()
    elif action == 'list_priorities':
        bot.list_priorities()
    elif action == 'create_form':
        if form_name == u'all':
            for fn in bot.form_def.keys():
                bot.create_form(fn, options.cut_off)
        else:
            bot.create_form(form_name, options.cut_off)
    elif action == 'create_categories':
        bot.create_category_pages()

if __name__ == '__main__':
    from optparse import OptionParser
    usage = "usage: %prog [options] username:password excel_file action\n"\
            "    excel_file -- FORSYS_WIKI_properties from Google Docs\n"\
            "    action -- push_properties, clean_properties, "\
            "list_priorities, create_form, create_categories\n"
    parser = OptionParser(usage)
    parser.add_option('-d', '--debug', dest='debug', action='store_true',
                      help='Run post mortem debug mode enabled')
    parser.add_option('-c', '--cutoff', dest='cut_off', action='store',
                  default=None,
                  help='priority score cut off value when creating DSS form')
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
    if action == 'create_form':
        try:
            form_name = unicode(args[3], 'utf8')
            form_name = form_name.replace(u'\xa0', ' ')
        except:
            parser.error('Give the form name when creating form')
    if not os.path.exists(excel):
        parser.error('The file %s does not exist' % excel)
    if action not in ACTIONS:
        parser.error('Action must be one of: %s' % ACTIONS)

    main(options, username, pwd, excel, action, form_name)
