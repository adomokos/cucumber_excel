Feature: Modify values in the Excel sheet
  In order to avoid silly mistakes
  As a user
  I want to interact with Excel
	
  Background:
    Given I have 2 categories
    And I have 3 child elements under the first category
	
  Scenario: Display column headers and captions
    Given the default values were loaded
    When I open the Excel workbook
    Then I should see "Category" in the "A1" cell
     And I should see "Value Total" in the "B1" cell
     And I should see "Value1" in the "C1" cell
     And I should see "Value2" in the "D1" cell
     And I should see "Category1" in the "A2" cell
     And I should see "Child1" in the "A3" cell
     And I should see "Child2" in the "A4" cell
	
  Scenario: Display loaded values
    Given the default values were loaded
    When I open the Excel workbook
    Then I should see 111 in the "C3" cell
     And I should see 353 in the "C2" cell
     And I should see 458 in the "B3" cell
     And I should see 1523 in the "B2" cell
	
  @focus
  Scenario Outline: Change values
    Given the default values were loaded
    When I open the Excel workbook
     And I put <child1_value> in the "C3" cell
    Then I should see <category_value1> in the "C2" cell
     And I should see <child1_sum> in the "B3" cell
     And I should see <category_total> in the "B2" cell
     And I should see <value1_total> in the "C9" cell
		
    Examples:
    | child1_value | category_value1 | child1_sum | category_total | value1_total |
    | 211          | 453             | 558        | 1623           | 1281         |
    | 51           | 293             | 398        | 1463           | 1121         |  