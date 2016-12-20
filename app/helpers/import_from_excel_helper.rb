module ImportFromExcelHelper
  
  def format_error(done,total,error_messages)
    final_message=""
    final_message=%{#{done} out of #{total} new issues have been created}
    final_message << %{<p>Details:</p>}
    final_message << %{<ul>}
    for message in error_messages
      final_message << %{<li>#{message}</li>}
    end
    final_message << %{</ul>}
    final_message << %{<p>#{l(:required_format)}</p>}
    return final_message
  end
end