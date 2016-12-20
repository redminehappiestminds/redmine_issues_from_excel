class ImportFromExcelController < ApplicationController  
  Mime::Type.register "application/xlsx", :xlsx
  
  #Mime::Type.register "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", :xlsx
  before_filter :get_project_data, :authorize ,:except =>[:download_sample]
  
  helper :import_from_excel
  include ImportFromExcelHelper
  include ApplicationHelper

  menu_item :issues
  
  require 'iconv'  
  require 'roo'
  
 def generate_excel()
    p = Axlsx::Package.new    
    wb = p.workbook  
    green_cell =  wb.styles.add_style(:bg_color => "008000", :fg_color => "FF", :sz => 10, :alignment => { :horizontal=> :center },:wrap_text => true)
    white_cell =  wb.styles.add_style(:fg_color => "FF", :sz => 20, :alignment => { :horizontal=> :center },:wrap_text => true)
    heading = ['Subject','Description','Estimation Time(hrs)', 'Start Date (yyyy-mm-dd)','Due Date (yyyy-mm-dd)']    
    list = []
    non_list = []
    data =[]
     if @custom_field_values.size > 0   
     @custom_field_values.each do |value| 
        if value.custom_field.is_required? and value.custom_field.field_format != 'list'        
          non_list.push(value.custom_field.name)
        end
        if value.custom_field.is_required? and value.custom_field.field_format == 'list'          
          list.push(value.custom_field.name)
          data +=value.custom_field.possible_values                   
        end
     end
     end
     column = heading.size+non_list.size       
    
    wb.add_worksheet(name: "sample") do |sheet|    
      data_column = 0
      sheet.add_row heading+non_list+list, :style=> green_cell             
        
      if @custom_field_values.size > 0   
     @custom_field_values.each do |value|       
        if value.custom_field.is_required? and value.custom_field.field_format == 'list'        
          data_list = value.custom_field.possible_values                   
          sheet.add_data_validation("#{Axlsx::cell_r(column,1)}:#{Axlsx::cell_r(column,100)}", {
          :type => :list,
          :formula1 => '"'+data_list.join(",")+'"',
          :showDropDown => false,
          :showErrorMessage => true,
          :errorTitle => '',
          :error => 'Please use the dropdown selector to choose',
          :errorStyle => :stop,
          :showInputMessage => true,
          :promptTitle => '',
          :prompt => ''})                 
          column = column+1   
          data_column += data_list.size     
        end 
       end
      end  
    
    end
    p_data = p.to_stream.read  
    return p_data
 end   
  
 def get_tracker(tracker_id)
   @project=  get_project_data()
   @issue = Issue.new
   @issue.project = @project
   @issue.tracker = @project.trackers.find(tracker_id || :first)
   @custom_field_values = @issue.editable_custom_field_values   
 end 
  
 def download_sample
   get_tracker(params[:tracker].to_i)  
    
   respond_to do |format|     
     format.xlsx do   
      send_data generate_excel(), type: "application/xlsx", filename: "#{@issue.tracker.name}.xlsx"
     end             
    end                                                                          
  end  
  
 def index
    respond_to do |format|
      format.html
    end
  end

  def get_issue_by_tracker(data, tracker)    
    col = Array.new
    data.each do |key,value|
      col << value
    end
    subject=col[0]
    description=col[1];estimated_hrs=col[2]
    start_date=col[3];due_date=col[4]
    issue = Issue.new
    issue.project = @project
    issue.author = User.current
    issue.tracker = tracker    
    issue.subject = subject
    issue.description = description
    issue.start_date= start_date
    issue.due_date=due_date
    issue.estimated_hours=estimated_hrs.to_f  unless estimated_hrs.blank?
    custom_field_hash = Hash.new
    if @custom_field_values.size > 0
      count = 4
      @custom_field_values.each do |value|
        if value.custom_field.is_required? and value.custom_field.field_format != 'list'        
          count = count +1
          custom_field_hash[value.custom_field.id] = col[count]
        end
      end
       @custom_field_values.each do |value|
        if value.custom_field.is_required? and value.custom_field.field_format == 'list'        
          count = count +1
          custom_field_hash[value.custom_field.id] = col[count]
        end
      end
    end  
        
    issue.custom_field_values =  custom_field_hash
    return issue
  end

  def open_spreadsheet(file)
    case File.extname(file.original_filename)    
    when '.ods' then Roo::OpenOffice.new(file.path, nil, :ignore)
    when '.xls' then Roo::Excel.new(file.path, nil, :ignore)
    when '.xlsx' then Roo::Excelx.new(file.path, nil, :ignore)
    else raise "Unknown file type: #{file.original_filename}"
    end
  end

  def excel_import
    if params[:dump][:file].blank?
      error = 'Please, Select Excel file'
      redirect_with_error error, @project    
    else
      begin        
        done = 0;total = 0
        max_row_to_read = Setting["plugin_redmine_issues_from_excel"]['max_row_to_read']
        error_messages = []
        Rails.cache.write("error_messages",[])
        tracker = @project.trackers.find(params[:dump][:tracker_id])
        spreadsheet = open_spreadsheet(params[:dump][:file])
        spreadsheet.default_sheet = spreadsheet.sheets.first 
        header = spreadsheet.row(1)
         get_tracker(params[:dump][:tracker_id].to_i)
       # (2..spreadsheet.last_row).each_with_index do |i, index|       
        (2..max_row_to_read).each_with_index do |i, index|         
         
          if not spreadsheet.row(i).join("").present?
            break
          end
          
          row = Hash[[header, spreadsheet.row(i)].transpose]
          # next if index == 0
          total = total+1          
          issue =get_issue_by_tracker(row, tracker)          
          if issue.save
          done=done+1
          else # invalid
            if issue.has_attribute? 'text_id'
              duplicate_issue = Issue.all(:conditions => "project_id = #{@project.id}")
              if duplicate_issue.blank?
                msg = Rails.cache.read("error_messages") << "Line:#{index+1}..Error: #{issue.errors.full_messages.uniq.join(', ')}"
              
                Rails.cache.write("error_messages", msg )
              else
                i_id = duplicate_issue.first.id
                msg = Rails.cache.read("error_messages") <<  "Line:#{index+1}..Error: #{issue.errors.full_messages.uniq.join(', ')} for this issue <a href=\"/issues/#{i_id}\">##{i_id}</a>"
              
                Rails.cache.write("error_messages", msg )
              end
            else
              msg = Rails.cache.read("error_messages") <<  "Line:#{index+1}..Error: #{issue.errors.full_messages.uniq.join(', ')}"
              
              Rails.cache.write("error_messages", msg)
            end
          end
        end
      rescue CSV::MalformedCSVError => e
        redirect_with_error e.message, @project
      return
      end      
      if done == total        
        flash[:notice]="Excel Import Successful, #{done} new issues have been created"
      else        
        errors = Rails.cache.read("error_messages")        
        Rails.cache.write("error_messages", format_error(done,total,errors))        
        redirect_to :controller=>"import_from_excel",:action=>"index",:project_id=>@project.identifier
        return
      end      
      redirect_to :controller=>"issues",:action=>"index",:project_id=>@project.identifier
    end
  end

  def redirect_with_error(err,project)
    flash[:error]=err
    redirect_to :controller=>"import_from_excel",:action=>"index",:project_id=>project.identifier
  end

  def get_project_data
    @project = Project.find(params[:project_id])
  rescue ActiveRecord::RecordNotFound
    render_404
    end
end
