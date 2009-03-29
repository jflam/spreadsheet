include System::Windows

# Run code with file-system operations pointed
# at the application's xap file, since the tests
# run in their own
def run_from_application
  tests = DynamicApplication.xap_file
  DynamicApplication.xap_file = nil
  yield
  DynamicApplication.xap_file = tests
end

run_from_application do
  require 'Spreadsheet'
  include Spreadsheet
end

$model = Application.current.root_visual.model

describe "Spreadsheet" do
  should 'return the value that we poked into the spreadsheet' do
    $model.set_cell "A1", "2"
    $model.get_cell("A1").should.equal "2".to_clr_string
  end
  
  should 'execute a formula' do
    $model.set_cell 'A1', '2'
    $model.set_cell 'A2', '=A1 + 2'
    $model.get_cell('A2').should.equal '4'.to_clr_string
  end
  
  should 'execute a formula and whitespace should not be significant' do
    $model.set_cell 'A1', '2'
    $model.set_cell 'A2', '=A1+2'
    $model.get_cell('A2').should.equal '4'.to_clr_string
  end
end

describe "CellParser" do
  should 'generate column names correctly' do
    CellParser.generate_column_name(0).should.equal 'A'.to_clr_string
    CellParser.generate_column_name(25).should.equal 'Z'.to_clr_string
    CellParser.generate_column_name(26).should.equal 'AA'.to_clr_string
    CellParser.generate_column_name(51).should.equal 'AZ'.to_clr_string
    CellParser.generate_column_name(52).should.equal 'BA'.to_clr_string
    CellParser.generate_column_name(701).should.equal 'ZZ'.to_clr_string
    # how handle exceptions in bacon?
  end
  
  should 'parse cell names correctly' do
    c = CellParser.parse_cell_name('A1')
    c.row.should.equal 0
    c.col.should.equal 'A'.to_clr_string
    c = CellParser.parse_cell_name('A99')
    c.row.should.equal 98
    c.col.should.equal 'A'.to_clr_string 
    c = CellParser.parse_cell_name('AA1')
    c.row.should.equal 0
    c.col.should.equal 'AA'.to_clr_string 
    c = CellParser.parse_cell_name('ZZ1')
    c.row.should.equal 0
    c.col.should.equal 'ZZ'.to_clr_string
  end
end
