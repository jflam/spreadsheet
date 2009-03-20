require 'code'

class ScriptController < Controller
  def save
    return 'Error: must be a POST request' unless self.request.http_method.to_s == "POST"
    
    Code.save self.request.Form["code"].to_s
    true
  end

  def index
    Code.get
  end
end
