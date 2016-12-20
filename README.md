# Import Issues from Excel

Redmine_import_from_Excel is a Redmine plugin to add the feature of importing issues From an Excel file.
For a team member with add issue permission, it shows a link 'Import issues from Excel' to the right on issues page.
This link redirects user to select excel file page.

## Installation

* Compatible with redmine 3 and Rails 4.
* Clons the plugin into the plugins directory

  ```
    cd redmine
    cd plugins/
    git clone git://github.com/redminehappiestminds/redmine_issues_from_excel.git
  ```
 
## Configuration
  
* You'll need to enable "Import issues from csv"  under "Issue tracking" section in Administration >  roles and permissions
* You can specify the Maximum Maximum Excel Rows to read in  Administration >  plugins > Redmine Issue From Excel plugin > Configure

## Help

* You can download sample excel file for reference using 'Sample Excel' link. 
   This samples file specifies all the required fields as heading in excel to create selected issue.

##Credits

