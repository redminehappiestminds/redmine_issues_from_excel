require 'redmine'

require_dependency 'import_from_excel_hooks'

Redmine::Plugin.register :redmine_issues_from_excel do
  name 'Redmine Issue From Excel plugin'
  author 'Happiest Minds Pvt Ltd'
  author_url 'http://www.happiestminds.com/'
  url 'https://github.com/redminehappiestminds/redmine_issues_from_excel.git'
  description 'This plugin for Redmine allows to import issues from excel file.'
  version '0.0.1'
  project_module :issue_tracking do
    permission :import_issues_from_excel, {:import_from_excel => [:index, :excel_import]} ,:require => :member
  end
  settings  :partial => 'settings/edit',
    :default => {
       "max_row_to_read" => 100
    }
    
end
