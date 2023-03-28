Feature: DataSnipper
  Scenario: User can do basic snip
    When Opening file "features\test.pdf"
    Then Result should be "For testing Document Upload"