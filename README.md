# Form-Duplicator
Form Duplicator is a tool to duplicate a public Google Form without being an owner/collaborator of the form.





Spreadsheet with script:

https://docs.google.com/spreadsheets/d/1nYBfvTkIWfmBFhnaNFrDDQh5m_9GJ0m0xHPbqVQhfD4/edit?usp=sharing

This script can copy the following Form question types:

- Short Answer
- Paragraph
- Multiple Choice
- Checkboxes
- Dropdown
- Linear Scale
- Panel
- Multiple choice grid
- Checkbox grid
- Section
- Date
- Time
- Image
- Video


File upload question type is not supported.

# Usage

## Worksheets

### Settings

Form to duplicate: The public form url you need to duplicate

Generated Form: After the script execution, it will display the duplicated form url (Editor page)

### Not duplicated

Some Form settings, question and question properties cannot be setted programmaticaly with Google Apps Script, because of the lack of methods / Object or due to open bugs.
If the script find some of these element will print every not duplicated settings of the form in this sheet.

Here is a list of things that cannot be duplicated from a Form with this tool:

#### Form settings

- Responses
  - Collect email addresses
    - Verified - method not present
    - Send responders a copy of their response
      - Always - method not present
      - When requested - method not present

- Presentation
  - Disable autosave for all respondents - method not present


#### Form Questions

In each question type it's impossible to set the text style (bold, italic ecc.) in the title and description.

Short Answer

- Response Validation
  - Text
    - Contains - method not present (work-around using Regular Expression literals)
    - Doesn't contain - method not present (work-around using Regular Expression literals)
 
Multiple Choice / Checkboxes

- Shuffle Order Question - method not present

Multiple Choice

- Due to a bug it's impossible to set the "Other" option if other choices in same question have "Go to section based on answer"
  See here: https://issuetracker.google.com/issues/171782147?pli=1

    Date
      - Due to a bug, using addDateTimeItem() will create a Date question without the time

    
    File Upload - Object not present

