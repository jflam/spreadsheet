class Code
  DB_FILE = File.dirname(__FILE__) + "/../db/code.dat"
  
  def self.get
    data = nil
    File.open(DB_FILE, "rb"){|f| data = f.read}
    Marshal.load(data)
  end

  def self.save(code)
    data = Marshal.dump(code)
    File.open(DB_FILE, "wb"){|f| f.write data}
  end
end
