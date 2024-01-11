/*
  Some Form settings cannot be setted programmaticaly with Google Apps Script.

  Form settings

    Responses

      Collect email addresses 
        Verified - method not present

      Send responders a copy of their response
        Always - method not present
        When requested - method not present

    Presentation
      Disable autosave for all respondents: method not present


  Form Questions

    In each question type it's impossible to set the text style (bold, italic ecc.) in the title, description ecc.

    Short Answer
      Response Validation
        Text
          Contains - method not present (work-around using Regular Expression literals)
          Doesn't contain - method not present (work-around using Regular Expression literals)

    Multiple Choice / Checkboxes:
      Shuffle Order Question - method not present

    Multiple Choice
      Due to a bug it's impossible to set the "Other" option if other choices in same question have "Go to section based on answer"
      See here: https://issuetracker.google.com/issues/171782147?pli=1

    Date
      Due to a bug, using addDateTimeItem() will create a Date question without the time

    
    File Upload - Object not present
  
*/

function duplicateForm() {
  const settingsSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = settingsSpreadsheet.getSheetByName("Settings");
  const formURL = settingsSheet.getRange("B3").getValue();
  const notDuplicatedSheet = settingsSpreadsheet.getSheetByName("Not duplicated");

  // Contains all the settings not duplicated
  var notDuplicated = [];

  // Contains all sections items
  var sections = {};




  /* 
    Contains question items and the possible choices for multiple choice / dopdown
    [[formQuestion, [questionChoices]]]

    When you call form.addQuestion() every new question is added to the last section added in the Form. For this reason, if a multiple choice / dropdown question has the go to section based on answer, I need to create the relative section before setting it as "Go to section based on answer". 
  */
  var answers = [];


  // Clear Errors worksheet
  notDuplicatedSheet.getRange("A2:B" + notDuplicatedSheet.getLastRow() + 1).clear();

  // Check if form url in the sheet is empty
  if (formURL == "") {
    SpreadsheetApp.getUi().alert("Insert a Google Form url");
    return;
  }

  var formPage;

  // Check Google Form url
  try {
    formPage = UrlFetchApp.fetch(formURL).getContentText();
  }
  catch (e) {
    SpreadsheetApp.getUi().alert("Google Form url not valid");
    return;
  }


  // All data of the form is contained in the FB_PUBLIC_LOAD_DATA_ variable in the form page source
  var tmp = formPage.split("FB_PUBLIC_LOAD_DATA_ = ")[1];
  var publicLoadData = JSON.parse(tmp.split(";</script>")[0]);


  // publicLoadData[3] is the Google Form file name
  var form = FormApp.create(publicLoadData[3]);

  form.setTitle(publicLoadData[1][8]);
  form.setDescription(publicLoadData[1][0])

  // Form settings
  if (publicLoadData[1][16] != null) {
    if (publicLoadData[1][16][2] == 1) {
      form.setIsQuiz(true);
    }
  }

  if (publicLoadData[1][2] != null) {

    form.setConfirmationMessage(publicLoadData[1][2][0]);

    if (publicLoadData[1][2][1] == 1) {
      form.setShowLinkToRespondAgain(true);
    }
    if (publicLoadData[1][2][2] == 1) {
      form.setPublishingSummary(true);
    }
    if (publicLoadData[1][2][3] == 1) {
      form.setAllowResponseEdits(true);
    }
  }

  if (publicLoadData[1][10] != null) {
    if (publicLoadData[1][10][1] == 1) {
      form.setLimitOneResponsePerUser(true);
    }
    if (publicLoadData[1][10][0] == 1) {
      form.setProgressBar(true);
    }
    if (publicLoadData[1][10][2] == 1) {
      form.setShuffleQuestions(true);
    }
    if (publicLoadData[1][10][6] == 3) {
      form.setCollectEmail(true);
    }


    // Not supported by Google Apps Script

    // Requrie login
    if (publicLoadData[1][10][6] == 2) {
      notDuplicated.push(["Form Setting", "Collect email addresses - Verified"]);
    }

    // Disable autosave for all respondents
    if (publicLoadData[1][10][5] == 1) {
      notDuplicated.push(["Form Setting", "Disable autosave for all respondents - Enabled"]);
    }

    /*  Send responders a copy of their response
        1 When requested
        2 Off
        3 Always 
    */

    if (publicLoadData[1][10][3] == 1) {
      notDuplicated.push(["Form Setting", "Send responders a copy of their response - When requested"]);
    }
    else if (publicLoadData[1][10][3] == 3) {
      notDuplicated.push(["Form Setting", "Send responders a copy of their response - Always"]);
    }
  }


  var questions = publicLoadData[1][1];


  /*
    Type of question

    0  Short Answer
    1  Paragraph
    2  Multiple Choice
    3  Dropdown
    4  Checkboxes
    5  Linear Scale
    6  Panel
    7  Multiple choice grid / Checkbox grid
    8  Section
    9  Date
    10 Time
    11 Image
    12 Video
    13 File Upload
  */

  /*
    Response Validation for Short Answer / Paragraph

      1 Number
        1 Greater than
        2 Grether than or equal to
        3 Less than
        4 Less than or equal to
        5 Equal to
        6 Not Equal to
        7 Between
        8 Not Between
        9 is number
        10 Whole number

      2 Text
        100 Contains
        101 Doesn't Contains
        102 Email
        103 URL

      4 Regular Expression
        299 Contains
        300 Doesn't Contains
        301 Matches
        302 Doesn't Matches

      6 Length
        202 Maximum character count
        203 Minimum character count
  */

  var item;
  var questionType;
  var validation;
  var choices;
  var choicesValues = [];

  if (questions != null) {
    for (let n = 0; n < questions.length; n++) {

      questionType = questions[n][3];

      switch (questionType) {
        case 0:
          item = form.addTextItem();
          item.setTitle(questions[n][1]);
          item.setHelpText(questions[n][2]);

          if (questions[n][4][0][2] == 1) {
            item.setRequired(true);
          }

          // Check if validation is present
          if (questions[n][4][0][4] != null) {

            switch (questions[n][4][0][4][0][1]) {
              case 1:
                validation = FormApp.createTextValidation()
                  .requireNumberGreaterThan(questions[n][4][0][4][0][2][0])
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
              case 2:
                validation = FormApp.createTextValidation()
                  .requireNumberGreaterThanOrEqualTo(questions[n][4][0][4][0][2][0])
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
              case 3:
                validation = FormApp.createTextValidation()
                  .requireNumberLessThan(questions[n][4][0][4][0][2][0])
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
              case 4:
                validation = FormApp.createTextValidation()
                  .requireNumberLessThanOrEqualTo(questions[n][4][0][4][0][2][0])
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
              case 5:
                validation = FormApp.createTextValidation()
                  .requireNumberEqualTo(questions[n][4][0][4][0][2][0])
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
              case 6:
                validation = FormApp.createTextValidation()
                  .requireNumberNotEqualTo(questions[n][4][0][4][0][2][0])
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
              case 7:
                validation = FormApp.createTextValidation()
                  .requireNumberBetween(questions[n][4][0][4][0][2][0], questions[n][4][0][4][0][2][1])
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
              case 8:
                validation = FormApp.createTextValidation()
                  .requireNumberNotBetween(questions[n][4][0][4][0][2][0], questions[n][4][0][4][0][2][1])
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
              case 9:
                validation = FormApp.createTextValidation()
                  .requireNumber()
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
              case 10:
                validation = FormApp.createTextValidation()
                  .requireWholeNumber()
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
              // There is no method to set "Contains / Doesn't Contain" validation, as a work-around I use Regular Expression 
              case 100:
                validation = FormApp.createTextValidation()
                  .requireTextContainsPattern(escapeRegex(questions[n][4][0][4][0][2][0]))
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
              case 101:
                validation = FormApp.createTextValidation()
                  .requireTextDoesNotContainPattern(escapeRegex(questions[n][4][0][4][0][2][0]))
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
              case 102:
                validation = FormApp.createTextValidation()
                  .requireTextIsEmail(questions[n][4][0][4][0][2][0])
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
              case 103:
                validation = FormApp.createTextValidation()
                  .requireTextIsUrl(questions[n][4][0][4][0][2][0])
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
              case 299:
                validation = FormApp.createTextValidation()
                  .requireTextContainsPattern(questions[n][4][0][4][0][2][0])
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
              case 300:
                validation = FormApp.createTextValidation()
                  .requireTextDoesNotContainPattern(questions[n][4][0][4][0][2][0])
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
              case 301:
                validation = FormApp.createTextValidation()
                  .requireTextMatchesPattern(questions[n][4][0][4][0][2][0])
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
              case 302:
                validation = FormApp.createTextValidation()
                  .requireTextDoesNotMatchPattern(questions[n][4][0][4][0][2][0])
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
              case 202:
                validation = FormApp.createTextValidation()
                  .requireTextLengthLessThanOrEqualTo(questions[n][4][0][4][0][2][0])
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
              case 203:
                validation = FormApp.createTextValidation()
                  .requireTextLengthGreaterThanOrEqualTo(questions[n][4][0][4][0][2][0])
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
            }
          }
          break;
        case 1:

          item = form.addParagraphTextItem();
          item.setTitle(questions[n][1]);
          item.setHelpText(questions[n][2]);

          if (questions[n][4][0][2] == 1) {
            item.setRequired(true);
          }

          if (questions[n][4][0][4] != null) {

            switch (questions[n][4][0][4][0][1]) {
              case 299:
                validation = validation = FormApp.createParagraphTextValidation()
                  .requireTextContainsPattern(questions[n][4][0][4][0][2][0])
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
              case 300:
                validation = FormApp.createParagraphTextValidation()
                  .requireTextDoesNotContainPattern(questions[n][4][0][4][0][2][0])
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
              case 301:
                validation = FormApp.createParagraphTextValidation()
                  .requireTextMatchesPattern(questions[n][4][0][4][0][2][0])
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
              case 302:
                validation = FormApp.createParagraphTextValidation()
                  .requireTextDoesNotMatchPattern(questions[n][4][0][4][0][2][0])
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
              case 202:
                validation = FormApp.createParagraphTextValidation()
                  .requireTextLengthLessThanOrEqualTo(questions[n][4][0][4][0][2][0])
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
              case 203:
                validation = FormApp.createParagraphTextValidation()
                  .requireTextLengthGreaterThanOrEqualTo(questions[n][4][0][4][0][2][0])
                  .setHelpText(questions[n][4][0][4][0][3])
                  .build();
                item.setValidation(validation);
                break;
            }
          }
          break;
        case 2:
          item = form.addMultipleChoiceItem();
          item.setTitle(questions[n][1]);
          item.setHelpText(questions[n][2]);

          if (questions[n][4][0][2] == 1) {
            item.setRequired(true);
          }

          // Method for Shuffle order options not present 
          if (questions[n][4][0][8] == 1) {
            notDuplicated.push([item.getTitle(), "Shuffle option Order - Enabled"]);
          }

          answers.push([item, questions[n][4][0][1]]);
          break;
        case 3:

          item = form.addListItem();
          item.setTitle(questions[n][1]);
          item.setHelpText(questions[n][2]);

          if (questions[n][4][0][2] == 1) {
            item.setRequired(true);
          }

          answers.push([item, questions[n][4][0][1]]);
          break;
        case 4:

          item = form.addCheckboxItem();
          item.setTitle(questions[n][1]);
          item.setHelpText(questions[n][2]);

          if (questions[n][4][0][2] == 1) {
            item.setRequired(true);
          }

          choices = questions[n][4][0][1];
          choicesValues = [];
          for (let i = 0; i < choices.length; i++) {
            choicesValues.push(item.createChoice(choices[i][0]));
          }

          item.setChoices(choicesValues)

          // response validation

          validation = questions[n][4][0][4];

          if (validation != null) {

            switch (validation[0][1]) {
              case 200:
                item.setValidation(FormApp.createCheckboxValidation().setHelpText(validation[0][3]).requireSelectAtLeast(validation[0][2]).build());
                break;
              case 201:
                item.setValidation(FormApp.createCheckboxValidation().setHelpText(validation[0][3]).requireSelectAtMost(validation[0][2]).build());
                break;
              case 204:
                item.setValidation(FormApp.createCheckboxValidation().setHelpText(validation[0][3]).requireSelectExactly(validation[0][2]).build());
                break;
            }
          }

          if (questions[n][4][0][8] == 1) {
            notDuplicated.push([item.getTitle(), "Shuffle option Order - Enabled"]);
          }

          break;
        case 5:

          item = form.addScaleItem();
          item.setTitle(questions[n][1]);
          item.setHelpText(questions[n][2]);

          if (questions[n][4][0][2] == 1) {
            item.setRequired(true);
          }

          choices = questions[n][4][0][1];
          item.setBounds(choices[0][0], choices[choices.length - 1][0]);

          item.setLabels(questions[n][4][0][3][0], questions[n][4][0][3][1]);
          break;
        case 6:
          item = form.addSectionHeaderItem();
          item.setTitle(questions[n][1]);
          item.setHelpText(questions[n][2]);
          break;
        case 7:
          var rows = [];
          if (questions[n][4][0][11][0] == 0) {
            item = form.addGridItem();
            if (questions[n][8] != null) {
              item.setValidation(FormApp.createGridValidation().setHelpText("Select one item per column.").requireLimitOneResponsePerColumn().build());
            }
          }
          else {
            item = form.addCheckboxGridItem();
            if (questions[n][8] != null) {
              item.setValidation(FormApp.createCheckboxGridValidation().setHelpText("Select one item per column.").requireLimitOneResponsePerColumn().build());
            }
          }
          item.setTitle(questions[n][1]);
          item.setHelpText(questions[n][2]);

          item.setColumns(questions[n][4][0][1]);

          for (let i = 0; i < questions[n][4].length; i++) {
            rows.push(questions[n][4][i][3][0]);
          }
          item.setRows(rows);

          break;
        case 8:
          item = form.addPageBreakItem();
          item.setTitle(questions[n][1]);
          item.setHelpText(questions[n][2]);
          sections[questions[n][0]] = [item, questions[n][5]];


          break;
        case 9:
          if (questions[n][4][0][7][0] == 0) {

            item = form.addDateItem();
          }
          else {
            // Doesn't create a Date with time but only Date
            item = form.addDateTimeItem();
          }
          if (questions[n][4][0][7][1] == 0) {
            item.setIncludesYear(false);
          }
          else {
            item.setIncludesYear(true);
          }

          item.setTitle(questions[n][1]);
          item.setHelpText(questions[n][2]);

          if (questions[n][4][0][2] == 1) {
            item.setRequired(true);
          }

          break;
        case 10:

          if (questions[n][4][0][6][0] == 0) {
            item = form.addTimeItem();
          }
          else {
            item = form.addDurationItem();
          }

          item.setTitle(questions[n][1]);
          item.setHelpText(questions[n][2]);

          if (questions[n][4][0][2] == 1) {
            item.setRequired(true);
          }

          break;
        case 11:

          item = form.addImageItem();
          item.setTitle(questions[n][1]);
          item.setHelpText(questions[n][2]);
          var alignment = questions[n][6][2][2];
          switch (alignment) {
            case 0:
              item.setAlignment(FormApp.Alignment.LEFT);
              break;
            case 1:
              item.setAlignment(FormApp.Alignment.CENTER);
              break;
            case 2:
              item.setAlignment(FormApp.Alignment.RIGHT);
              break;
          }
          break;
        case 12:
          item = form.addVideoItem();
          item.setTitle(questions[n][1]);
          item.setHelpText(questions[n][2]);
          item.setVideoUrl("https://www.youtube.com/watch?v=" + questions[n][6][3]);
          var alignment = questions[n][6][2][2];
          switch (alignment) {
            case 0:
              item.setAlignment(FormApp.Alignment.LEFT);
              break;
            case 1:
              item.setAlignment(FormApp.Alignment.CENTER);
              break;
            case 2:
              item.setAlignment(FormApp.Alignment.RIGHT);
              break;
          }
          break;
        case 13:
          // File upload not supported yet in Apps Script
          notDuplicated.push([questions[n][1], "Wasn't created"]);
          break;

      }
    }


    // For multiple choice / dropdown
    var question;
    var hasOtherOption = false;
    var hasGoToSections = false;
    var goToSection;

    for (let n = 0; n < answers.length; n++) {

      choicesValues = [];
      hasOtherOption = false;
      hasGoToSections = false;
      question = answers[n][0];
      choices = answers[n][1];

      for (let i = 0; i < choices.length; i++) {

        goToSection = choices[i][2];

        // Check if question has the "Other" option
        if (choices[i][4] == 1) {
          hasOtherOption = true;
        }
        else {
          if (goToSection != null) {
            hasGoToSections = true;
            switch (goToSection) {
              case -1:
                choicesValues.push(question.createChoice(choices[i][0], FormApp.PageNavigationType.RESTART));
                break;
              case -2:
                choicesValues.push(question.createChoice(choices[i][0], FormApp.PageNavigationType.CONTINUE));
                break;
              case -3:
                choicesValues.push(question.createChoice(choices[i][0], FormApp.PageNavigationType.SUBMIT));
                break;
              default:
                choicesValues.push(question.createChoice(choices[i][0], sections[goToSection][0]));
                break;
            }
          }
          else {
            choicesValues.push(question.createChoice(choices[i][0]));
          }
        }

      }

      /* 
        If question has "Other" option enabled and other choices have Go to section based on answer enabled,
        doesn't create the other option (Google Apps Script bug: https://issuetracker.google.com/issues/171782147?pli=1)
      */
      if (hasOtherOption) {
        if (!hasGoToSections) {
          question.showOtherOption(true);
        }
        else {
          notDuplicated.push([question.getTitle(), "Other option - Enabled"]);
        }
      }

      question.setChoices(choicesValues);
    }





    // For sections
    var section;

    for (const [key, value] of Object.entries(sections)) {

      section = value[0];
      goToSection = value[1];

      // Check if has go to section based on answer
      if (goToSection != null) {
        switch (goToSection) {
          case -1:
            section.setGoToPage(FormApp.PageNavigationType.RESTART);
            break;
          case -2:
            section.setGoToPage(FormApp.PageNavigationType.CONTINUE);
            break;
          case -3:
            section.setGoToPage(FormApp.PageNavigationType.SUBMIT);
            break;
          default:
            section.setGoToPage(sections[goToSection][0]);
            break;
        }
      }
    }
  }


  settingsSheet.getRange("B5").setValue(form.getEditUrl());

  if (notDuplicated.length > 0) {
    notDuplicatedSheet.getRange("A2:B" + (notDuplicated.length + 1)).setValues(notDuplicated);
    SpreadsheetApp.getUi().alert("Form partially duplicated. See the Not duplicated worksheet for more information");
  }
  else {
    SpreadsheetApp.getUi().alert("Form duplicated correctly");
  }
}



function escapeRegex(string) {
  return string.replace(/[/\-\\^$*+?.()|[\]{}]/g, '\\$&');
}
