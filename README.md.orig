semanticbot
===========

Semantic data integration between FORSYS Google Docs documents and FORSYS semantic mediawiki

General
-------

The raison d'être for this script is to map the forsys_semanticwiki_properties document (http://tinyurl.com/amwtpz3) to a Semantic Mediawiki implementation. The document in question describes the categories, properties and forms the wiki implementation should have.

The structure of the document has evolved organically over the FORSYS COST action lifetime, and can seem rather convoluted in it's structure for the first time viewer. Initially it was used to describe just one category ("DSS description") and it's properties, and the structure of a single form for that category. This is still reflected in its structure

The 'special' DSS form
----------------------

The DSS description form is still special in how its structure is defined on the document. The document has different sections (see column E "Title"). These are the sections the DSS description form is divided into. The other forms are not divided into separate sections (and hence into separate templates in the wiki). The order of the properties on the form is dictated by the section and document order for the DSS form (section order hardcoded in the script). For other forms either a dedicated column is used to give the order (see column C "FormCase" and colunm D "FormLesson"), or it's based purely on the order of the properties in the document.

What ends up in what form
-------------------------

Column B "Target form" lists the categories (and hence forms, a category having a default edit form) the property is used for. Given as semicolon separated list of category names

Default edit form for a category
--------------------------------

For the "Edit with form" tab to appear for a page belonging to a category, the category page needs a [[Has default form::<category name>]] property on its page, e.g. [[Has default form::DSS]] for the DSS category to use the DSS form as the default form. This is not taken care by the script, so this must be set manually after creating the forms.

Properties of type Page
-----------------------

If the type of a property is Page, and column I "Related category for property" has a value, the value of the property is a link to a page belonging to that category.

The page can be globally shared between the linking category pages created with the form in which case column J "Type of related category" has value "General". In this case the script will create a page for each value given in the column K "Value" and these pages are assigned to the category defined in column I "Related category for property". This is done with the command create_categories command of the script.

If column J "Type of related category" is "Specific", each page created with the form will link to its own page belonging to the category given in column I "Related category for property", i.e. the related category pages are not shared between the "main" category pages, but each get's its own. In this case the property will have a default value linking to a as of yet nonexisting page of the related category.


Form control
------------

Column O "Form control" is used to define what kind of UI control is used for the property on the form. First comes the control name (if any), a semicolon and the size of the control.

The size parameter is the width of the field for the text field (default control) and combobox, and the number of values shown for listbox.

If the column J "Type of related category" has the value "Specific", and the "Form control" has value "default page", the UI text field will have a default value derived from the name of the page being edited with the form and the name of the property.

Conditional fields
------------------

Some of the properties are intended to be shown on the form only if another property has a specific value. This is controlled with the column R "Depends on". First the id of the "master" property is given, and separated with a semicolon, the value this property should have in order for the "detail" property to be shown. If the "detail" property is assigned to several forms, this "master"-"detail" relationship is only effective if both properties are on the form.


