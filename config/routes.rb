#custom routes for this plugin

if Rails::VERSION::MAJOR >= 3
  RedmineApp::Application.routes.draw do
   scope "/projects/:project_id" , :name_prefix => 'project' do 
     resources :import_from_excel do 
       collection do
         post :excel_import
         get :download_sample
       end
    end
   end
  end
else
 ActionController::Routing::Routes.draw do |map|
  map.resources :import_from_excel, :name_prefix => 'project_', :path_prefix => '/projects/:project_id',:collection => {:excel_import => :post}    
 end 
end
