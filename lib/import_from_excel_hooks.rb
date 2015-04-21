class ImportFromExcelHooks < Redmine::Hook::ViewListener
  render_on :view_issues_sidebar_issues_bottom, :partial => 'issues/import_issues_from_excel'
end
