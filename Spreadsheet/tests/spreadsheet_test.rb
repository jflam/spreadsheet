include System::Windows

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

require 'Spreadsheet'
include Spreadsheet

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
    c.col.should.equal 0
    c = CellParser.parse_cell_name('A99')
    c.row.should.equal 98
    c.col.should.equal 0
    c = CellParser.parse_cell_name('AA1')
    c.row.should.equal 0
    c.col.should.equal 26
    c = CellParser.parse_cell_name('ZZ1')
    c.row.should.equal 0
    c.col.should.equal 701
  end
end