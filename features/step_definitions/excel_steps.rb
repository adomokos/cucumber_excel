Given /^I have 2 categories$/ do
  # setting up data - ignore it in this test
end

Given /^I have 3 child elements under the first category$/ do
  # setting up data - ignore it in this test
end

Given /^the default values were loaded$/ do
  # setting up data - ignore it in this test
end

When /^I open the Excel workbook$/ do
  ExcelHandler.instance.open_excel
  @worksheet = ExcelHandler.instance.worksheet
  @worksheet.extend CellValueGetter
end

When /^I put (\d+) in the "([^\"]*)" cell$/ do |value, cell|
  @worksheet.set_cell_value(cell, value)
end

Then /^I should see "([^\"]*)" in the "([^\"]*)" cell$/ do |value, cell|
  @worksheet.get_cell_value(cell).strip.should == value	
end

Then /^I should see (\d+) in the "([^\"]*)" cell$/ do |value, cell|
  @worksheet.get_cell_value(cell).should == value.to_i
end

at_exit do
  ExcelHandler.instance.close_excel
end

module CellValueGetter
  def get_cell_value(cell)
    get_cell(cell).value
  end
	
  def set_cell_value(cell, value)
    get_cell(cell).value = value
  end
	
  def get_cell(cell)
    cell_values = cell.split('')
    cell_values.length.should == 2
    cells(cell_values[1].to_i, cell_values[0])
  end
	
end