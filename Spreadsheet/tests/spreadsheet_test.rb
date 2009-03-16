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

