var scriptTitle = "formMule Script V6.5.3 (1/30/14)";
var scriptName = 'formMule'
var scriptTrackingId = 'UA-30976195-1'
// Written by Andrew Stillman for New Visions for Public Schools
// Published under GNU General Public License, version 3 (GPL-3.0)
// See restrictions at http://www.opensource.org/licenses/gpl-3.0.html
// Support and contact at http://www.youpd.org/formmule
// 

var ss = SpreadsheetApp.getActiveSpreadsheet();
var MULEICONURL = 'https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/formMule6_icon.gif?attachauth=ANoY7cpzkYpzCxNW5Xfg-odHtcod8xUm5hP7l1Q7JgEFVbem83zgoz0mAq40-26SBZM7LlpODEluBzq6-0j9I51YFHcV1ex-DYen2UyYg7mXfuSJkcqJFE5T9yD9La27lk1Wh3oBwjeozxJLUQIuFWPd2dSTSs_eFF8v-t4EYrEJ1bJjifRHLalWrZQUilFs9HjkNWh_1x7IhGgqNWhZdDN_PAUZn1Dd0niCuUtX4gTcFl6obSZ-dFBuCHB0IEg3TjEVtKtZbKGj&attredirects=0';

function onInstall () {
  Browser.msgBox('To complete initialization, please select \"Run initial installation\" from the formMule script menu above');
  var menuEntries = [];
      menuEntries.push({name: "What is formMule?", functionName: "formMule_whatIs"});
      menuEntries.push({name: "Run initial installation", functionName: "formMule_completeInstall"});
  onOpen();
}

function onOpen() {
  var menuEntries = [];
  var installed = ScriptProperties.getProperty('installedFlag');
  var sheetName = ScriptProperties.getProperty('sheetName');
  var webAppUrl = ScriptProperties.getProperty('webAppUrl');
  var twilioNumber = ScriptProperties.getProperty('twilioNumber');
  if (!(installed)) {
      menuEntries.push({name: "Run initial installation", functionName: "formMule_completeInstall"});
  } else {
      menuEntries.push({name: "What is formMule?", functionName: "formMule_whatIs"});
      menuEntries.push({name: "Step 1: Define merge source settings", functionName: "formMule_defineSettings"});
    if ((sheetName) && (sheetName!='')) {
      menuEntries.push({name: "Step 2a: Set up email merge", functionName: "formMule_emailSettings"});
      menuEntries.push({name: "Step 2b: Set up calendar merge", functionName: "formMule_setCalendarSettings"});
      if (!webAppUrl) {
        menuEntries.push({name: "Step 2c: Set up SMS and Voice Message merge", functionName: "formMule_howToPublishAsWebApp"});
      }
      if ((webAppUrl)&&(!twilioNumber)) {
        menuEntries.push({name: "Step 2c: Set up SMS and Voice Message merge", functionName: "formMule_howToSetUpTwilio"});
      }
      if ((webAppUrl)&&(twilioNumber)) {
        menuEntries.push({name: "Step 2c: Set up SMS and Voice Message merge", functionName: "formMule_smsAndVoiceSettings"});
      }     
      menuEntries.push({name: "Preview and perform manual merge", functionName: "formMule_loadingPreview"});
      menuEntries.push({name: "Advanced options", functionName: "formMule_advanced"});
    }
  }
  this.ss.addMenu("formMule", menuEntries);
}

function formMule_advanced() {
  var app = UiApp.createApplication().setTitle("Advanced options").setHeight(130).setWidth(290);
  var quitHandler = app.createServerHandler('formMule_quitUi');
  var handler1 = app.createServerHandler('formMule_pasteFormUrl');
  var button1 = app.createButton('Copy down formulas on form submit').addClickHandler(quitHandler).addClickHandler(handler1);
  var handler2 = app.createServerHandler('formMule_extractorWindow');
  var button2 = app.createButton('Package this workflow for others to copy').addClickHandler(quitHandler).addClickHandler(handler2);
  var handler3 = app.createServerHandler('formMule_institutionalTrackingUi');
  var button3 = app.createButton('Manage your usage tracker settings').addClickHandler(quitHandler).addClickHandler(handler3);
  var panel = app.createVerticalPanel();
  panel.add(button1);
  panel.add(button2);
  panel.add(button3);
  app.add(panel);
  ss.show(app);
  return app;
}

function formMule_quitUi(e) {
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}

function formMule_pasteFormUrl(){
  var app = UiApp.createApplication().setTitle("We Need Your Form URL!");
  var panel = app.createVerticalPanel().setSpacing(15)
  var label = app.createLabel("Please copy and paste the live url to your form into this box:");
  var inputBox = app.createTextArea().setName('formUrl').setWidth("300px").setHeight("150px")
  var handler = app.createServerHandler('saveFormUrlAndDetectFormSheet').addCallbackElement(inputBox);
  var quitHandler = app.createServerHandler('formMule_quitUi');
  var button = app.createButton("Save").addClickHandler(quitHandler).addClickHandler(handler);
  panel.add(label).add(inputBox).add(button)
  app.add(panel)
  ss.show(app)
  return app  
}

function saveFormUrlAndDetectFormSheet(e){
  var app = UiApp.getActiveApplication();
  var formUrl = e.parameter.formUrl;
  ScriptProperties.setProperty('formUrl', formUrl);
  formMule_detectFormSheet()
}


function formMule_completeInstall() {
  formMule_preconfig();
  ScriptProperties.setProperty('installedFlag', 'true');
  var triggers = ScriptApp.getScriptTriggers();
  var formTriggerSetFlag = false;
  var editTriggerSetFlag = false;
  for (var i = 0; i<triggers.length; i++) {
    var eventType = triggers[i].getEventType();
    var triggerSource = triggers[i].getTriggerSource();
    var handlerFunction = triggers[i].getHandlerFunction();
    if ((handlerFunction=='formMule_onFormSubmit')&&(eventType=="ON_FORM_SUBMIT")&&(triggerSource=="SPREADSHEETS")) {
      formTriggerSetFlag = true;
    }
    if ((handlerFunction=='formMule_checkForSourceChanges')&&(eventType=="ON_EDIT")&&(triggerSource=="SPREADSHEETS")) {
      editTriggerSetFlag = true;
    }
  }
  if (formTriggerSetFlag==false) {
    formMule_setFormTrigger();
  }
  if (editTriggerSetFlag==false) {
    formMule_setEditTrigger();
  }
//ensure console and readme sheets exist

 onOpen();
}

function formMule_manualSend () {
  var lock = LockService.getPublicLock();
  lock.waitLock(120000);
  var manual = true;
  formMule_sendEmailsAndSetAppointments(manual);
  lock.releaseLock();
}


function formMule_setFormTrigger() {
  var ssKey = SpreadsheetApp.getActiveSpreadsheet().getId();
  ScriptApp.newTrigger('formMule_onFormSubmit').forSpreadsheet(ssKey).onFormSubmit().create();
}

function formMule_setEditTrigger() {
  var ssKey = SpreadsheetApp.getActiveSpreadsheet().getId();
  ScriptApp.newTrigger('formMule_checkForSourceChanges').forSpreadsheet(ssKey).onEdit().create();
}

function formMule_lockFormulaRow() {
// setFrozenRows function was once broken in Apps Script...seeing if it's fixed.
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheetName = ScriptProperties.getProperty('sheetName');
var sheet = ss.getSheetByName(sheetName);
sheet.setFrozenRows(2);
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheetName = ScriptProperties.getProperty('sheetName');
var sheet = ss.getSheetByName(sheetName)
ss.setActiveSheet(sheet);
  var frozenRows = sheet.getFrozenRows();
  if (frozenRows!=2) {
    Browser.msgBox("To avoid issues, it is highly recommended you freeze the first two rows of this sheet.");
  }
}


function formMule_setCalendarSettings () {
  var calendarStatus = ScriptProperties.getProperty('calendarStatus');
  var calendarUpdateStatus = ScriptProperties.getProperty('calendarUpdateStatus');
  var calendarToken  = ScriptProperties.getProperty('calendarToken');
  var eventTitleToken = ScriptProperties.getProperty('eventTitleToken');
  var locationToken = ScriptProperties.getProperty('locationToken');
  var guests = ScriptProperties.getProperty('guests');
  var emailInvites = ScriptProperties.getProperty('emailInvites');
  var allDay = ScriptProperties.getProperty('allDay');
  var startTimeToken = ScriptProperties.getProperty('startTimeToken');
  var endTimeToken = ScriptProperties.getProperty('endTimeToken');
  var descToken = ScriptProperties.getProperty('descToken');
  var reminderType = ScriptProperties.getProperty('reminderType');
  var minBefore = ScriptProperties.getProperty('minBefore');
 
  var app = UiApp.createApplication().setTitle('Step 2b: Set up calendar merge').setHeight(450).setWidth(800);
  var panel = app.createHorizontalPanel().setId('calendarPanel').setSpacing(20).setStyleAttribute('borderColor', 'grey');
 // var helpLabel = app.createLabel().setText('Indicate whether a calendar event should be generated on form submission, and (optionally) set an additional condition below. Fields for calendar id, and date fields can be populated with static values or dynamically using the variables to the right.');
 // var popupPanel = app.createPopupPanel();
 // popupPanel.add(helpLabel);
  var scrollPanel = app.createScrollPanel().setHeight("230px");
  var verticalPanel = app.createVerticalPanel();
  var topSettingsGrid = app.createGrid(4, 2).setId('topSettingsGrid');
  
  // Create calendar event conditions grid
  var conditionsGrid = app.createGrid(2,3).setId('conditionsGrid').setCellPadding(0);
  var conditionLabel = app.createLabel('Create Event Condition');
  var dropdown = app.createListBox().setId('col-0').setName('col-0').setWidth("150px").setEnabled(false);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheetName = ScriptProperties.getProperty('sheetName');
  if ((sourceSheetName)&&(sourceSheetName!='')) {
     var sourceSheet = ss.getSheetByName(sourceSheetName);
  } else {
    Browser.msgBox('You must select a source sheet before this menu item can be selected.');
    formMule_defineSettings();
    app().close();
    return app;
  }
  var lastCol = sourceSheet.getLastColumn();
  if (lastCol > 0) {
  var headers = sourceSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  } else {
    Browser.msgBox('You must have headers in your source data sheet before this menu item can be selected.');
    formMule_defineSettings();
    app().close();
    return app;
  }
  for (var i=0; i<headers.length; i++) {
    dropdown.addItem(headers[i]);
  }
  var equalsLabel = app.createLabel('=');
  var textbox = app.createTextBox().setId('val-0').setName('val-0').setEnabled(false);
  var conditionHelp = app.createLabel('Leave blank to always create an event. Use NULL for empty.  NOT NULL for not empty.').setStyleAttribute('fontSize', '10px');
  conditionsGrid.setWidget(0, 0, conditionLabel);
  conditionsGrid.setWidget(1, 0, dropdown);
  conditionsGrid.setWidget(1, 1, equalsLabel);
  conditionsGrid.setWidget(1, 2, textbox);
  conditionsGrid.setWidget(0, 2, conditionHelp);
  
  var extraHelp1 = app.createLabel('Note: Events are created only for rows where the condition is met AND \"Event Creation Status\" is blank').setStyleAttribute('fontSize', '10px').setStyleAttribute('color', '#0099FF').setStyleAttribute('fontWeight', 'bold').setId('extraHelp1').setVisible(false);
  var extraHelp2 = app.createLabel('Note: Events are updated or deleted only for rows where the condition is met AND \"Event Update Status\" is blank').setStyleAttribute('fontSize', '10px').setStyleAttribute('color', '#66C285').setStyleAttribute('fontWeight', 'bold').setId('extraHelp2').setVisible(false);
 
  var createPanel = app.createVerticalPanel();
  createPanel.add(conditionsGrid);
  createPanel.add(extraHelp1);
  
  var calendarConditionString = ScriptProperties.getProperty('calendarConditions');
  if ((calendarConditionString)&&(calendarConditionString!='')) {
    var calendarConditionsObject = Utilities.jsonParse(calendarConditionString);
    var selectedHeader = calendarConditionsObject['col-0'];
    var selectedIndex = 0;
    for (var i=0; i<headers.length; i++) {
      if (headers[i]==selectedHeader) {
        selectedIndex = i;
        break;
      }
    }
    dropdown.setSelectedIndex(selectedIndex);
    var selectedValue = calendarConditionsObject['val-0'];
    textbox.setValue(selectedValue);
  }
  
  //Update calendar conditions grid
  var updateConditionsGrid = app.createGrid(4,3).setId('updateConditionsGrid');
  var updateConditionLabel = app.createLabel('Update Event Condition');
  var deleteConditionLabel = app.createLabel('Delete Event Condition');
  var updateDropdown = app.createListBox().setId('update-col-0').setName('update-col-0').setWidth("150px").setEnabled(false);
  for (var i=0; i<headers.length; i++) {   
    updateDropdown.addItem(headers[i]);
  }
  
  var deleteDropdown = app.createListBox().setId('delete-col-0').setName('delete-col-0').setWidth("150px").setEnabled(false);
  for (var i=0; i<headers.length; i++) {
    deleteDropdown.addItem(headers[i]);
  }
  var equalsLabel = app.createLabel('if');
  var equalsLabel2 = app.createLabel('=');
  var updateTextbox = app.createTextBox().setId('update-val-0').setName('update-val-0').setEnabled(false);
  var deleteTextbox = app.createTextBox().setId('delete-val-0').setName('delete-val-0').setEnabled(false);
  var conditionHelp = app.createLabel('Leave blank to update no events. Use \"NULL\" for empty.  \"NOT NULL\" for not empty.').setStyleAttribute('fontSize', '10px');
  var deleteConditionHelp = app.createLabel('Leave blank to delete no events. Use \"NULL\" for empty.  \"NOT NULL\" for not empty.').setStyleAttribute('fontSize', '10px');
  updateConditionsGrid.setWidget(0, 0, updateConditionLabel);
  updateConditionsGrid.setWidget(1, 0, updateDropdown);
  updateConditionsGrid.setWidget(1, 1, equalsLabel);
  updateConditionsGrid.setWidget(1, 2, updateTextbox);
  updateConditionsGrid.setWidget(2, 0, deleteConditionLabel);
  updateConditionsGrid.setWidget(3, 0, deleteDropdown);
  updateConditionsGrid.setWidget(3, 1, equalsLabel2);
  updateConditionsGrid.setWidget(3, 2, deleteTextbox);
  updateConditionsGrid.setWidget(0, 2, conditionHelp);
  updateConditionsGrid.setWidget(2, 2, deleteConditionHelp);
  
  var calendarUpdateConditionString = ScriptProperties.getProperty('calendarUpdateConditions');
  if ((calendarUpdateConditionString)&&(calendarUpdateConditionString!='')) {
    var calendarUpdateConditionsObject = Utilities.jsonParse(calendarUpdateConditionString);
    var selectedUpdateHeader = calendarUpdateConditionsObject['col-0'];
    var selectedIndex = 0;
    for (var i=0; i<headers.length; i++) {
      if (headers[i]==selectedUpdateHeader) {
        selectedIndex = i;
        break;
      }
    }
    updateDropdown.setSelectedIndex(selectedIndex);
    var selectedValue = calendarUpdateConditionsObject['val-0'];
    updateTextbox.setValue(selectedValue);
  }
  
  var calendarDeleteConditionString = ScriptProperties.getProperty('calendarDeleteConditions');
  if ((calendarDeleteConditionString)&&(calendarDeleteConditionString!='')) {
    var calendarDeleteConditionsObject = Utilities.jsonParse(calendarDeleteConditionString);
    var selectedDeleteHeader = calendarDeleteConditionsObject['col-0'];
    var selectedIndex = 0;
    for (var i=0; i<headers.length; i++) {
      if (headers[i]==selectedDeleteHeader) {
        selectedIndex = i;
        break;
      }
    }
    deleteDropdown.setSelectedIndex(selectedIndex);
    var selectedValue = calendarDeleteConditionsObject['val-0'];
    deleteTextbox.setValue(selectedValue);
  }
  

  var eventIdPanel = app.createVerticalPanel();
  var eventIdLabel = app.createLabel('Column containing Event Id to be used for update');
  var eventIdCol = app.createListBox().setId('eventIdCol').setName('eventIdCol').setWidth("150px").setEnabled(false);
  for (var i=0; i<headers.length; i++) {
    eventIdCol.addItem(headers[i]);
  } 
  var selectedEventIdHeader = ScriptProperties.getProperty('eventIdCol');
  var eventIdExists = headers.indexOf("Event Id");
  if (eventIdExists==-1) {
    eventIdCol.addItem('Event Id');
  } 
  if ((!(selectedEventIdHeader))||(selectedEventIdHeader=='')) {
    selectedEventIdHeader = "Event Id";
  }
  
  var selectedEventIdUpdateIndex = 0;
  for (var i=0; i<headers.length; i++) {
      if (headers[i]==selectedEventIdHeader) {
        selectedEventIdUpdateIndex = i;
        break;
      }
    }
  eventIdCol.setSelectedIndex(selectedEventIdUpdateIndex);
  
  
  eventIdPanel.add(eventIdLabel);
  eventIdPanel.add(eventIdCol);
  
  var settingsGrid = app.createGrid(14, 2).setId('settingsGrid').setWidth(520);
  var calendarLabel = app.createLabel().setText('Calendar Id (xyz@sample.org)');
  var calendarTextBox = app.createTextBox().setName('calendarToken').setWidth("100%");
  if (calendarToken) { calendarTextBox.setValue(calendarToken); }
  var eventTitleLabel = app.createLabel().setText('Event title');
  var eventTitleTextBox = app.createTextBox().setName('eventTitleToken').setWidth("100%");
  if (eventTitleToken) { eventTitleTextBox.setValue(eventTitleToken); }
  var eventLocationLabel = app.createLabel().setText('Location');
  var eventLocationTextBox = app.createTextBox().setName('locationToken').setWidth("100%");
  if (locationToken) { eventLocationTextBox.setValue(locationToken); }
  var guestsLabel = app.createLabel().setText('Guests (comma separated email addresses)');
  var guestsTextBox = app.createTextBox().setName('guests').setWidth("100%");
  if (guests) { guestsTextBox.setValue(guests);  }
  var emailInvitesCheckBox = app.createCheckBox().setText('Email invitations').setName('emailInvites');
  if (emailInvites=="true") { emailInvitesCheckBox.setValue(true); }
  var startTimeLabel = app.createLabel().setText('Start time (must use a spreadsheet formatted datetime value)');
  var startTimeTextBox = app.createTextBox().setName('startTimeToken').setWidth("100%");
  if (startTimeToken) { startTimeTextBox.setValue(startTimeToken); }
  var endTimeLabel = app.createLabel().setText('End time (must use a spreadsheet formatted datetime value)');
  var endTimeTextBox = app.createTextBox().setName('endTimeToken').setWidth("100%");
  if (endTimeToken) { endTimeTextBox.setValue(endTimeToken); }
  var allDayCheckBox = app.createCheckBox().setText('All day event').setId('allDayTrue').setName('allDay').setVisible(false);
  var allDayCheckBoxFalse = app.createCheckBox().setText('All day event').setId('allDayFalse').setVisible(true);
  if (allDay=="true") { 
    allDayCheckBox.setValue(true).setVisible(true);
    allDayCheckBoxFalse.setVisible(false);
    endTimeLabel.setVisible(false);
    endTimeTextBox.setVisible(false);
   }
  var uncheckClientHandler = app.createClientHandler().forTargets(endTimeLabel, endTimeTextBox).setVisible(true);
                                                                                                           
  var uncheckServerHandler = app.createServerHandler('formMule_uncheck').addCallbackElement(allDayCheckBox);
  
  var checkClientHandler = app.createClientHandler().forTargets(endTimeLabel, endTimeTextBox).setVisible(false);                                                   
 
  var checkServerHandler = app.createServerHandler('formMule_check').addCallbackElement(allDayCheckBox);
  
  allDayCheckBox.addClickHandler(uncheckClientHandler).addClickHandler(uncheckServerHandler);
  allDayCheckBoxFalse.addClickHandler(checkClientHandler).addClickHandler(checkServerHandler);

  var descLabel = app.createLabel().setText('Event description  (HTML accepted, however tags will unfortunately get stripped upon next edit of these settings.)');
  var descTextArea = app.createTextArea().setName('descToken').setWidth("100%").setHeight(75);
  if (descToken) { descTextArea.setValue(descToken); }
  var reminderLabel = app.createLabel().setText('Set reminder type');
  var reminderListBox = app.createListBox().setName('reminderType');
  reminderListBox.addItem('None');
  reminderListBox.addItem('Email reminder');
  reminderListBox.addItem('Popup reminder');
  reminderListBox.addItem('SMS reminder');
   if (reminderType) { 
    switch(reminderType) {
    case "Email reminder": 
       reminderListBox.setSelectedIndex(1);
    break;
    case "Popup reminder":
       reminderListBox.setSelectedIndex(2);
    case "SMS reminder":
       reminderListBox.setSelectedIndex(3);
    break;
    default:
        reminderListBox.setSelectedIndex(0);
    }
  }
  
  var reminderMinLabel = app.createLabel().setText('Minutes before');
  var reminderMinTextBox = app.createTextBox().setName('minBefore');
  if (minBefore) { 
    reminderMinTextBox.setValue(minBefore);
  }
  var repeatLabel = app.createLabel('Repeat for how many weeks (leave blank for single event)');
  var repeatTextBox = app.createTextBox().setName('calendarWeeklyRepeats');
  var calendarWeeklyRepeats = ScriptProperties.getProperty('calendarWeeklyRepeats');
  if (calendarWeeklyRepeats) {
    repeatTextBox.setValue(calendarWeeklyRepeats)
  }
  var daysLabel = app.createLabel('Comma separated days of the week to repeat the event on (e.g. Monday,Wednesday)');
  var daysBox = app.createTextBox().setName('calendarWeekdays');
  var calendarWeekdays = ScriptProperties.getProperty('calendarWeekdays');
  if (calendarWeekdays) {
    daysBox.setValue(calendarWeekdays);
  }
  if (minBefore) { reminderMinTextBox.setValue(minBefore); }
  settingsGrid.setWidget(0, 0, calendarLabel).setWidget(0, 1, calendarTextBox)
              .setWidget(1, 0, eventTitleLabel).setWidget(1, 1, eventTitleTextBox)
              .setWidget(2, 0, eventLocationLabel).setWidget(2, 1, eventLocationTextBox) 
              .setWidget(3, 0, guestsLabel).setWidget(3, 1, guestsTextBox) 
              .setWidget(4, 0, emailInvitesCheckBox)
              .setWidget(5, 0, allDayCheckBox)
              .setWidget(6, 0, allDayCheckBoxFalse)
              .setWidget(7, 0, startTimeLabel).setWidget(7, 1, startTimeTextBox)
              .setWidget(8, 0, endTimeLabel).setWidget(8, 1, endTimeTextBox)
              .setWidget(9, 0, descLabel).setWidget(9, 1, descTextArea)
              .setWidget(10, 0, reminderLabel).setWidget(10, 1, reminderListBox)
              .setWidget(11, 0, reminderMinLabel).setWidget(11, 1, reminderMinTextBox)
              .setWidget(12, 0, repeatLabel)
              .setWidget(12, 1, repeatTextBox)
              .setWidget(13, 0, daysLabel)
              .setWidget(13, 1, daysBox);
  
  settingsGrid.setStyleAttribute("backgroundColor", "whiteSmoke").setStyleAttribute('textAlign', 'right').setStyleAttribute('padding', '8px');
  settingsGrid.setStyleAttribute(0, 1, 'width', '200px');
  
  
  
  var checkBox = app.createCheckBox().setText("Turn on calendar-event merge feature.").setId('calendarStatus').setName('calendarStatus').setStyleAttribute('color', '#0099FF').setStyleAttribute('fontWeight', 'bold');
 
  var clickHandler1 = app.createServerHandler('toggle1').addCallbackElement(topSettingsGrid);
  checkBox.addValueChangeHandler(clickHandler1);
      
  var calTrigger = ScriptProperties.getProperty('calTrigger');
  var calTriggerCheckBox = app.createCheckBox().setText("Trigger event creation on form submit.").setId('calTriggerCheckBox').setName('calTrigger').setStyleAttribute('color', '#0099FF').setStyleAttribute('fontWeight', 'bold').setEnabled(false);
  if (calTrigger=="true") {
    calTriggerCheckBox.setValue(true);
  }
  if (calendarStatus=="true") {
    checkBox.setValue(true);
    calTriggerCheckBox.setEnabled(true);
    dropdown.setEnabled(true);
    textbox.setEnabled(true);
    eventIdCol.setEnabled(true);
    extraHelp1.setVisible(true);
  }
 
       
  var updateCheckBox = app.createCheckBox().setText("Turn on calendar-event update feature.").setId('calendarUpdateStatus').setName('calendarUpdateStatus').setStyleAttribute('color', '#66C285').setStyleAttribute('fontWeight', 'bold');
 
  
  var clickHandler2 = app.createServerHandler('toggle2').addCallbackElement(topSettingsGrid);
  updateCheckBox.addValueChangeHandler(clickHandler2);
  
  var calUpdateTrigger = ScriptProperties.getProperty('calUpdateTrigger');
  var calUpdateTriggerCheckBox = app.createCheckBox().setText("Trigger event update on form submit.").setId('calUpdateTriggerCheckBox').setName('calUpdateTrigger').setStyleAttribute('color', '#66C285').setStyleAttribute('fontWeight', 'bold').setEnabled(false);
  if (calUpdateTrigger=="true") {
    calUpdateTriggerCheckBox.setValue(true);
  }
  
  if (calendarUpdateStatus=="true") {
    updateCheckBox.setValue(true);
    eventIdCol.setEnabled(true);
    calUpdateTriggerCheckBox.setEnabled(true);
    updateDropdown.setEnabled(true);
    updateTextbox.setEnabled(true);
    deleteDropdown.setEnabled(true);
    deleteTextbox.setEnabled(true);
    extraHelp2.setVisible(true);
  }
  
  
  //app.add(popupPanel);
  topSettingsGrid.setWidget(0, 0, checkBox);
  topSettingsGrid.setWidget(1, 0, calTriggerCheckBox);
  topSettingsGrid.setWidget(2, 0, createPanel);
  topSettingsGrid.setWidget(3, 0, extraHelp2);
  topSettingsGrid.setWidget(0, 1, updateCheckBox);
  topSettingsGrid.setWidget(1, 1, calUpdateTriggerCheckBox);
  topSettingsGrid.setWidget(2, 1, updateConditionsGrid);
  topSettingsGrid.setWidget(3, 1, eventIdPanel); 
  var verticalPanel = app.createVerticalPanel();
  var variablesPanel = app.createVerticalPanel().setStyleAttribute('backgroundColor', 'whiteSmoke').setStyleAttribute('padding', '10px');
  var variablesScrollPanel = app.createScrollPanel().setHeight(240);
  var variablesLabel = app.createLabel().setText("Choose from the following variables: ").setStyleAttribute('fontWeight', 'bold');
  var tags = formMule_getAvailableTags();
  var flexTable = app.createFlexTable()
  for (var i = 0; i<tags.length; i++) {
    var tag = app.createLabel().setText(tags[i]);
    flexTable.setWidget(i, 0, tag);
    }
  variablesPanel.add(variablesLabel);
  variablesScrollPanel.add(flexTable);
  variablesPanel.add(variablesScrollPanel);
  panel.add(settingsGrid);
  panel.add(variablesPanel);
  verticalPanel.add(panel);
  var mainPanel = app.createVerticalPanel();
  mainPanel.add(topSettingsGrid);
  var spinner = app.createImage(MULEICONURL).setWidth(150);
  spinner.setVisible(false);
  spinner.setStyleAttribute("position", "absolute");
  spinner.setStyleAttribute("top", "120px");
  spinner.setStyleAttribute("left", "220px");
  spinner.setId("dialogspinner");
  var submitHandler = app.createServerHandler('formMule_saveEmailSettings').addCallbackElement(topSettingsGrid).addCallbackElement(panel);
  var backgroundHandler = app.createClientHandler().forTargets(mainPanel, settingsGrid).setStyleAttribute('opacity', '0.5');
  var spinnerHandler = app.createClientHandler().forTargets(spinner).setVisible(true);
  var buttonHandler = app.createServerHandler('formMule_saveCalendarSettings').addCallbackElement(scrollPanel).addCallbackElement(topSettingsGrid);
  var button = app.createButton().setText('Save Calendar Settings').addClickHandler(buttonHandler);
  button.addMouseDownHandler(backgroundHandler).addMouseDownHandler(spinnerHandler);
  verticalPanel.add(button);
  scrollPanel.add(verticalPanel);
  mainPanel.add(scrollPanel);
  app.add(mainPanel);
  app.add(spinner);
  ss.show(app);
  return app;
}

function toggle1(e) {
  var app = UiApp.getActiveApplication();
  var calTriggerCheckBox = app.getElementById('calTriggerCheckBox');
  var textBox = app.getElementById('val-0');
  var dropDown = app.getElementById('col-0');
  var extraHelp1 = app.getElementById('extraHelp1');
  var calendarStatus = e.parameter.calendarStatus;
  if (calendarStatus == "true") {
    calTriggerCheckBox.setEnabled(true);
    textBox.setEnabled(true);
    dropDown.setEnabled(true);
    extraHelp1.setVisible(true);
  } else {
    calTriggerCheckBox.setValue(false);
    calTriggerCheckBox.setEnabled(false);
    textBox.setEnabled(false);
    dropDown.setEnabled(false);
    extraHelp1.setVisible(false);
  }
  return app;
}


function toggle2(e) {
  var app = UiApp.getActiveApplication();
  var calTriggerCheckBox = app.getElementById('calUpdateTriggerCheckBox');
  var deleteTextBox = app.getElementById('delete-val-0');
  var deleteDropDown = app.getElementById('delete-col-0');
  var textBox = app.getElementById('update-val-0');
  var dropDown = app.getElementById('update-col-0');
  var extraHelp2 = app.getElementById('extraHelp2');
  var eventIdCol = app.getElementById('eventIdCol');
  var calendarUpdateStatus = e.parameter.calendarUpdateStatus;
  if (calendarUpdateStatus == "true") {
    calTriggerCheckBox.setEnabled(true);
    textBox.setEnabled(true);
    dropDown.setEnabled(true);
    deleteTextBox.setEnabled(true);
    deleteDropDown.setEnabled(true);
    eventIdCol.setEnabled(true);
    extraHelp2.setVisible(true);
  } else {
    calTriggerCheckBox.setValue(false);
    calTriggerCheckBox.setEnabled(false);
    textBox.setEnabled(false);
    dropDown.setEnabled(false);
    deleteTextBox.setEnabled(false);
    deleteDropDown.setEnabled(false);
    eventIdCol.setEnabled(false);
    extraHelp2.setVisible(false);
  }
  return app;
}





function formMule_getAvailableTags() {
  var sheetName = ScriptProperties.getProperty("sheetName");
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Browser.msgBox('You must select a source data sheet before this menu item can be selected.');
    formMule_defineSettings();
  }
  var lastCol = sheet.getLastColumn();
  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  var headers = headerRange.getValues();
  var availableTags = [];
  
  for (var i=0; i<headers[0].length; i++) {
    availableTags[i] = "${\""+headers[0][i]+"\"}";
  } 
  var k = availableTags.length;
  availableTags[k] = "$currDay  (current day)";
  availableTags[k+1] = "$currMonth  (current month)";
  availableTags[k+2] = "$currYear  (current year)";
  availableTags[k+3] = "$eventId  (available in description only)";
  availableTags[k+4] = "$formUrl  (link to form)"
  return availableTags;
}


function formMule_check() {
//  Browser.msgBox("hellow");
  var app = UiApp.getActiveApplication();
  var allDayCheckBox = app.getElementById('allDayTrue');
  allDayCheckBox.setValue(true).setVisible(true);
  var allDayCheckBoxFalse = app.getElementById('allDayFalse');
  allDayCheckBoxFalse.setValue(false).setVisible(false);
  return app;
}

function formMule_uncheck() {
//  Browser.msgBox("hello");
  var app = UiApp.getActiveApplication();
  var allDayCheckBox = app.getElementById('allDayTrue');
  allDayCheckBox.setValue(false).setVisible(false);
  var allDayCheckBoxFalse = app.getElementById('allDayFalse');
  allDayCheckBoxFalse.setValue(false).setVisible(true);
  return app;
}

function formMule_fetchHeaders(sheet) {
  if (sheet.getLastColumn() < 1) {
    sheet.getRange(1, 1, 1, 3).setValues([['Dummy Header 1','Dummy Header 2','Dummy Header 3']]);
    SpreadsheetApp.flush();
  }
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers;  
}


function formMule_saveCalendarSettings (e) {
  var app = UiApp.getActiveApplication();
  var calendarStatus = e.parameter.calendarStatus;
  ScriptProperties.setProperty('calendarStatus', calendarStatus);
  var calendarUpdateStatus = e.parameter.calendarUpdateStatus;
  ScriptProperties.setProperty('calendarUpdateStatus', calendarUpdateStatus);
  var calTrigger = e.parameter.calTrigger;
  ScriptProperties.setProperty('calTrigger', calTrigger);
  var calUpdateTrigger = e.parameter.calUpdateTrigger;
  ScriptProperties.setProperty('calUpdateTrigger', calUpdateTrigger);
  var conditionObject = new Object();
  conditionObject['col-0'] = e.parameter['col-0'];
  conditionObject['val-0'] = e.parameter['val-0'].trim();
  conditionObject['sht-0'] = "Event Creation";
  var conditionString = Utilities.jsonStringify(conditionObject);
  ScriptProperties.setProperty('calendarConditions', conditionString);
  var updateConditionObject = new Object();
  updateConditionObject['col-0'] = e.parameter['update-col-0'];
  updateConditionObject['val-0'] = e.parameter['update-val-0'].trim();
  updateConditionObject['sht-0'] = "Event Update";
  var updateConditionString = Utilities.jsonStringify(updateConditionObject);
  ScriptProperties.setProperty('calendarUpdateConditions', updateConditionString);
  var deleteConditionObject = new Object();
  deleteConditionObject['col-0'] = e.parameter['delete-col-0'];
  deleteConditionObject['val-0'] = e.parameter['delete-val-0'].trim();
  deleteConditionObject['sht-0'] = "Event Update";
  var deleteConditionString = Utilities.jsonStringify(deleteConditionObject);
  ScriptProperties.setProperty('calendarDeleteConditions', deleteConditionString);
  
  //check for illegal choices
  if ((e.parameter.calendarStatus=="true")&&(e.parameter.calendarUpdateStatus=="true")&&(e.parameter['col-0']==e.parameter['update-col-0'])&&(e.parameter['val-0']==e.parameter['update-val-0'])&&(e.parameter['update-val-0']!='')) {
    Browser.msgBox("You cannot set identical conditions for calendar event creation and calendar event update, as this would create and update an event at the same time.");
    app.close();
    formMule_setCalendarSettings();
    return app;
  }
  if ((e.parameter.calendarStatus=="true")&&(e.parameter.calendarUpdateStatus=="true")&&(e.parameter['col-0']==e.parameter['delete-col-0'])&&(e.parameter['val-0']==e.parameter['delete-val-0'])&&(e.parameter['delete-val-0']!='')) {
    Browser.msgBox("You cannot set identical conditions for calendar event creation and calendar event deletion, as this would create and delete an event at the same time.");   
    app.close();
    formMule_setCalendarSettings();
    return app;
  }
  
  if ((e.parameter.calendarUpdateStatus=="true")&&(e.parameter['update-col-0']==e.parameter['delete-col-0'])&&(e.parameter['update-val-0']==e.parameter['delete-val-0'])&&(e.parameter['delete-val-0']!='')) {
    Browser.msgBox("You cannot set identical conditions for calendar event update and calendar event deletion, as this would update and delete an event at the same time.");   
    app.close();
    formMule_setCalendarSettings();
    return app; 
  }  
  
  var eventIdCol = e.parameter.eventIdCol;
  ScriptProperties.setProperty('eventIdCol', eventIdCol);
  var sheetName = ScriptProperties.getProperty('sheetName');
  var sheet = ss.getSheetByName(sheetName);
  var headers = formMule_fetchHeaders(sheet);
  if (calendarStatus=="true") {
    headers = formMule_fetchHeaders(sheet);
    var eventIdIndex = headers.indexOf("Event Id");
    var statusIndex = headers.indexOf("Event Creation Status");
    if((statusIndex==-1)&&(eventIdIndex==-1)) {
      var lastCol = sheet.getLastColumn();
      sheet.insertColumnAfter(lastCol);
      sheet.getRange(1, lastCol+1).setValue('Event Id').setBackground("#0099FF").setFontColor("black").setComment("Don't change the name of this column.");
      var copyDownOption = ScriptProperties.getProperty('copyDownOption');
      if (copyDownOption=="true") {
        sheet.getRange(2, lastCol+1).setValue('N/A: This is the formula row.').setBackground("#0099FF").setFontColor("black");
      }
      sheet.insertColumnAfter(lastCol+1);
      sheet.getRange(1, lastCol+2).setValue('Event Creation Status').setBackground("#0099FF").setFontColor("black").setComment("Don't change the name of this column.");
      var copyDownOption = ScriptProperties.getProperty('copyDownOption');
      if (copyDownOption=="true") {
        sheet.getRange(2, lastCol+2).setValue('N/A: This is the formula row.').setBackground("#0099FF").setFontColor("black");
      }
    }
  }
    if (calendarUpdateStatus=="true") {
    headers = formMule_fetchHeaders(sheet);
    var eventUpdateIndex = headers.indexOf("Event Update Status");
    var eventIdIndex = headers.indexOf("Event Id");
    if((eventIdIndex!=-1)&&(eventUpdateIndex==-1)) {
      var lastCol = sheet.getLastColumn();
      sheet.insertColumnAfter(lastCol);
      sheet.getRange(1, lastCol+1).setValue('Event Update Status').setBackground("#66C285").setFontColor("black").setComment("Don't change the name of this column.");
      var copyDownOption = ScriptProperties.getProperty('copyDownOption');
      if (copyDownOption=="true") {
        sheet.getRange(2, lastCol+1).setValue('N/A: This is the formula row.').setBackground("#66C285").setFontColor("black");
      }
    }
  }
    var calendarToken = e.parameter.calendarToken;
    var eventTitleToken = e.parameter.eventTitleToken;
    var locationToken = e.parameter.locationToken;
    var allDay = e.parameter.allDay;
    var startTimeToken = e.parameter.startTimeToken;
    var endTimeToken = e.parameter.endTimeToken;
    var descToken = e.parameter.descToken;
    var guests = e.parameter.guests;
    var emailInvites = e.parameter.emailInvites;
    var reminderType = e.parameter.reminderType;
    var sendInvites = e.parameter.sendInvites;
    var minBefore = e.parameter.minBefore;
    var calendarWeeklyRepeats = e.parameter.calendarWeeklyRepeats;
    var calendarWeekdays = e.parameter.calendarWeekdays;
  
    ScriptProperties.setProperty('calendarToken', calendarToken);
    ScriptProperties.setProperty('eventTitleToken', eventTitleToken);
    ScriptProperties.setProperty('locationToken', locationToken);
    ScriptProperties.setProperty('guests', guests);
    ScriptProperties.setProperty('emailInvites', emailInvites);
    ScriptProperties.setProperty('allDay', allDay);
    ScriptProperties.setProperty('startTimeToken', startTimeToken);
    ScriptProperties.setProperty('endTimeToken', endTimeToken);
    ScriptProperties.setProperty('descToken', descToken);
    ScriptProperties.setProperty('reminderType', reminderType);
    ScriptProperties.setProperty('minBefore', minBefore);
    ScriptProperties.setProperty('calendarWeeklyRepeats', calendarWeeklyRepeats);
    ScriptProperties.setProperty('calendarWeekdays', calendarWeekdays);
  
    var errMsg = '';
    if (calendarToken=='') { errMsg += "You forgot to enter a Calendar Id, "; }
    if (eventTitleToken=='') { errMsg += "You forgot to enter an event title, "; }
    if ((reminderType!="None")&&(minBefore=='')) { errMsg += "You forgot to specify the numer of minutes before the event for reminders, "; }
    if ((allDay!='true')&&(startTimeToken=='')) { errMsg += "You forgot to enter a start time, "; }
    if ((allDay!='true')&&(endTimeToken=='')) { errMsg += "You forgot to enter an end time, "; }
    if (errMsg !='') {
      Browser.msgBox(errMsg);
      formMule_setCalendarSettings ();
      app.close();
      return app;
    }
   app.close();
   return app;
}
  

function formMule_onFormSubmit () {
  var properties = ScriptProperties.getProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
     ss = SpreadsheetApp.openById(properties.ssId);
  }
  var submissionRow = ss.getActiveRange().getRow();
  var lock = LockService.getPublicLock();
  lock.waitLock(120000);
  var sheetName = ScriptProperties.getProperty('sheetName');
  var sheet = ss.getSheetByName(sheetName);
  var headers = formMule_fetchHeaders(sheet);
  var caseNoSetting = ScriptProperties.getProperty('caseNoSetting');
  if (caseNoSetting == "true") {
    var headers = formMule_fetchHeaders(sheet);
    var caseNoIndex = headers.indexOf("Case No");
    var cellRange = sheet.getRange(submissionRow, caseNoIndex+1);
    var cellValue = cellRange.getValue();
    if (cellValue=="") {
      cellRange.setValue(formMule_assignCaseNo());
    }
  }
  copyDownFormulas(submissionRow, properties);
  formMule_sendEmailsAndSetAppointments();
  lock.releaseLock();
  debugger;
}


function urlencode(inNum) {
  // Function to convert non URL compatible characters to URL-encoded characters
  var outNum = 0;     // this will hold the answer
  outNum = escape(inNum); //this will URL Encode the value of inNum replacing whitespaces with %20, etc.
  return outNum;  // return the answer to the cell which has the formula
}




function formMule_assignCaseNo(resetToValue) {
  if (resetToValue) {
  var caseNo = resetToValue;
  } else { 
  var caseNo = ScriptProperties.getProperty('caseNo')
  };
  if(caseNo==null) {
      ScriptProperties.setProperty('caseNo','0');
      caseNo = 0;
   } else {
   caseNo = parseInt(caseNo) + 1;
   ScriptProperties.setProperty('caseNo', caseNo);
  }
return caseNo; 
}


function formMule_emailSettings () {
  var app = UiApp.createApplication().setTitle("Step 2a: Set up email merge").setHeight(450).setWidth(700);
  var panel = app.createVerticalPanel().setId("emailPanel").setStyleAttribute('backgroundColor', 'whiteSmoke').setStyleAttribute('padding', '5px').setStyleAttribute('marginTop', '5px');
  var scrollPanel = app.createScrollPanel().setHeight("150px");
  var helpLabel = app.createLabel().setText('Indicate below whether you want templated emails to be generated upon form submit, how many different emails you want to create, and (optionally) any additional conditions that must be met before sending');
  var helpPopup = app.createPopupPanel();
  helpPopup.add(helpLabel);
  app.add(helpPopup);
  
 
  var emailStatus = ScriptProperties.getProperty('emailStatus');
  var emailStatusCheckBox = app.createCheckBox().setText('Turn-on email merge feature').setName('emailStatus').setStyleAttribute('color', '#FF9999').setStyleAttribute('fontWeight','bold');
  if (emailStatus=="true") {
    emailStatusCheckBox.setValue(true);
  }
  var emailTrigger = ScriptProperties.getProperty('emailTrigger');
  var emailTriggerCheckBox = app.createCheckBox().setText('Trigger this feature on form submit').setName('emailTrigger').setStyleAttribute('color', '#FF9999').setStyleAttribute('fontWeight','bold');
  if (emailTrigger=="true") {
    emailTriggerCheckBox.setValue(true);
  }
  var numSelectLabel = app.createLabel().setText("How many different unique, templated emails do you want to send out to different recipients when the form is submitted?");
  var grid = app.createGrid().setId('emailConditionGrid');
  formMule_setEmailConditionGrid(app);
  var numSelectChangeHandler = app.createServerHandler('formMule_refreshEmailConditions').addCallbackElement(panel);
  var numSelect = app.createListBox().setId("numSelect").setName("numSelectValue").addChangeHandler(numSelectChangeHandler);
  numSelect.addItem('1');
  numSelect.addItem('2');
  numSelect.addItem('3');
  numSelect.addItem('4');
  numSelect.addItem('5');
  numSelect.addItem('6');
  var preSelectedNum = ScriptProperties.getProperty('numSelected');
  switch (preSelectedNum) {
    case "1":
     numSelect.setSelectedIndex(0);
    break;
    case "2":
     numSelect.setSelectedIndex(1);
    break;
    case "3":
     numSelect.setSelectedIndex(2);
    break;
     case "4":
     numSelect.setSelectedIndex(3);
    break;
     case "5":
     numSelect.setSelectedIndex(4);
    break;
    case "6":
     numSelect.setSelectedIndex(5);
    break;
    default: 
     numSelect.setSelectedIndex(0);
  }
  
  var spinner = app.createImage(MULEICONURL).setWidth(150);
  spinner.setVisible(false);
  spinner.setStyleAttribute("position", "absolute");
  spinner.setStyleAttribute("top", "120px");
  spinner.setStyleAttribute("left", "220px");
  spinner.setId("dialogspinner");
  
  
  var submitHandler = app.createServerHandler('formMule_saveEmailSettings').addCallbackElement(panel);
  var spinnerHandler = app.createClientHandler().forTargets(spinner).setVisible(true).forTargets(panel).setStyleAttribute('opacity', '0.5');
  var button = app.createButton().setText("Submit settings").addClickHandler(submitHandler).addMouseDownHandler(spinnerHandler);
  var extraHelp1 = app.createLabel('Note: Emails are generated only for rows where the condition is met AND the \"Send Status\" for that template is blank').setStyleAttribute('fontSize', '10px').setStyleAttribute('color', '#FF9999').setStyleAttribute('fontWeight', 'bold').setId('extraHelp1');
  var numSelectNote = app.createLabel().setStyleAttribute('marginTop', '20px').setText("The the first time you submit this form, it will auto-generate a new sheet with a blank template for each email. You will need to go to these sheets and complete the the \"To:\", \"CC:\", \"Subject:\", and \"Body:\" sections of the template. If you return later and increase the number here, the script will generate additional sheets. Decreasing this value will delete the corresponding sheets.  Do not change the names of the template sheets.");
  panel.add(emailStatusCheckBox);
  panel.add(emailTriggerCheckBox);
  panel.add(numSelectLabel);
  panel.add(numSelect);
  scrollPanel.add(grid);
  panel.add(scrollPanel);
  panel.add(extraHelp1);
  panel.add(numSelectNote); 
  panel.add(button);
  app.add(panel);
  app.add(spinner);
  this.ss.show(app);
}

function formMule_setEmailConditionGrid(app) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheetName = ScriptProperties.getProperty('sheetName');
  if ((!sourceSheetName)||(sourceSheetName=='')) {
    Browser.msgBox('You must select a source data sheet before this menu item can be selected.');
    formMule_defineSettings();
    app().close();
    return app;
  } else {
    var sourceSheet = ss.getSheetByName(sourceSheetName);
  }
  var lastCol = sourceSheet.getLastColumn();
  var headers = sourceSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var numSelected = ScriptProperties.getProperty('numSelected');
  if ((!numSelected)||(numSelected=='')) {
    numSelected = 1;
  }
  numSelected = parseInt(numSelected);
  var grid = app.getElementById('emailConditionGrid');
  grid.resize(numSelected+1, 5);
  grid.setWidget(0,0,app.createLabel('Send Email Template'));
  grid.setWidget(0,2,app.createLabel('Column'));
  grid.setWidget(0,4,app.createLabel('Leave blank to send for all rows. Use NULL for empty.  NOT NULL for not empty.').setStyleAttribute('fontSize', '10px'));
  var dropdown = [];
  var textbox = [];
  var namebox = [];
  for (var i=0; i<numSelected; i++) {
    var label = 'Condition for Email ' + (i+1);
    grid.setWidget(i+1, 0, app.createLabel(label)).setStyleAttribute(i+1, 0, 'width','100px');
    dropdown[i] = app.createListBox().setId('conditionCol-'+i).setName('conditionCol-'+i).setWidth("150px");
    for (var j=0; j<headers.length; j++) {
        dropdown[i].addItem(headers[j]);
    }
    namebox[i] = app.createTextBox().setId('templateName-'+i).setName('templateName-'+i);
    textbox[i] = app.createTextBox().setId('value-'+i).setName('value-'+i);
    grid.setWidget(i+1, 0, namebox[i]);
    grid.setWidget(i+1, 1, app.createLabel('if'));
    grid.setWidget(i+1, 2, dropdown[i]);
    grid.setWidget(i+1, 3, app.createLabel('='));
    grid.setWidget(i+1, 4, textbox[i]);
  }
  var emailConditions = ScriptProperties.getProperty('emailConditions');
  if ((emailConditions)&&(emailConditions!='')) {
    emailConditions = Utilities.jsonParse(emailConditions);
    numSelected = parseInt(numSelected);
    for (var i=0; i<numSelected; i++) {
      var preset = 0;
      for (var j=0; j<headers.length; j++) {
        if (emailConditions["col-"+i]==headers[j]) {
          preset = j;
          break;
        }
      }
      dropdown[i].setSelectedIndex(preset);
      textbox[i].setValue(emailConditions["val-"+i]);
      if ((emailConditions["sht-"+i])&&(emailConditions["sht-"+i]!='')) {
        namebox[i].setValue(emailConditions["sht-"+i]);
      } else {
        namebox[i].setValue("Email"+(i+1)+" Template");
      }
    }
  }
  
  if (!(emailConditions)) {
    namebox[0].setValue("Email1 Template");
  }
  return app; 
}


function formMule_refreshEmailConditions(e) {
  var app = UiApp.getActiveApplication();
  var grid = app.getElementById('emailConditionGrid');
  var numSelected = e.parameter.numSelectValue;
  var oldNumSelected = ScriptProperties.getProperty('numSelected');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheetName = ScriptProperties.getProperty('sheetName');
  if ((!sourceSheetName)||(sourceSheetName=='')) {
    var sourceSheet = ss.getSheets()[0];
  } else {
    var sourceSheet = ss.getSheetByName(sourceSheetName);
  }
  var lastCol = sourceSheet.getLastColumn();
  var headers = sourceSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  if ((!numSelected)||(numSelected=='')) {
    numSelected = 1;
  }
  numSelected = parseInt(numSelected);
  grid.resize(numSelected+1, 5);
  grid.setWidget(0,0,app.createLabel('Send Email Template'));
  grid.setWidget(0,2,app.createLabel('Column'));
  grid.setWidget(0,4,app.createLabel('Leave blank to ignore. Use NULL for empty.  NOT NULL for not empty.').setStyleAttribute('fontSize', '10px'));
  var dropdown = [];
  var namebox = [];
  var textbox = [];
  for (var i=0; i<numSelected; i++) {
    var label = 'Condition for Email ' + (i+1);
    grid.setWidget(i+1, 0, app.createLabel(label)).setStyleAttribute(i+1, 0, 'width','100px');
    dropdown[i] = app.createListBox().setId('conditionCol-'+i).setName('conditionCol-'+i).setWidth("150px");
    for (var j=0; j<lastCol; j++) {
        dropdown[i].addItem(headers[j]);
    }
    namebox[i] = app.createTextBox().setId('templateName-'+i).setName('templateName-'+i);
    textbox[i] = app.createTextBox().setId('value-'+i).setName('value-'+i);
    grid.setWidget(i+1, 0, namebox[i]);
    grid.setWidget(i+1, 1, app.createLabel('if'));
    grid.setWidget(i+1, 2, dropdown[i]);
    grid.setWidget(i+1, 3, app.createLabel('='));
    grid.setWidget(i+1, 4, textbox[i]);
  }
  var emailConditions = new Object();
  for (var i=0; i<numSelected; i++) {
    var condCol = e.parameter["conditionCol-"+i];
    var condVal = e.parameter["value-"+i];
    var shtName = e.parameter["templateName-"+i];
    emailConditions["col-"+i] = condCol;
    emailConditions["val-"+i] = condVal;
    emailConditions["sht-"+i] = shtName;
  }
  emailConditions = Utilities.jsonStringify(emailConditions);
  if ((emailConditions)&&(emailConditions!='')) {
    emailConditions = Utilities.jsonParse(emailConditions);
    if (numSelected>=oldNumSelected) {
    for (var i=0; i<oldNumSelected; i++) {    
      var preset = 0;
      for (var j=0; j<headers.length; j++) {
        if (emailConditions["col-"+i]==headers[j]) {
          preset = j;
          break;
        }
      }
      dropdown[i].setSelectedIndex(preset);
      textbox[i].setValue(emailConditions["val-"+i]);
     }
   }
   if (numSelected>=oldNumSelected) {
     for (var i=0; i<numSelected; i++) {    
       if ((emailConditions["sht-"+i]&&emailConditions["sht-"+i]!='')) {
        namebox[i].setValue(emailConditions["sht-"+i]);
      } else {
        namebox[i].setValue("Email"+(i+1)+" Template");
      }
     }
    }
    if (numSelected<oldNumSelected) {
    for (var i=0; i<numSelected; i++) {    
      var preset = 0;
      for (var j=0; j<headers.length; j++) {
        if (emailConditions["col-"+i]==headers[j]) {
          preset = j;
          break;
        }
      }
      dropdown[i].setSelectedIndex(preset);
      textbox[i].setValue(emailConditions["val-"+i]);
      if ((emailConditions["sht-"+i]&&emailConditions["sht-"+i]!='')) {
        namebox[i].setValue(emailConditions["sht-"+i]);
      } else {
        namebox[i].setValue("Email"+(i+1)+" Template");
      }
     }
   }
  }
  return app; 
}


//returns true if testval meets the condition 
function formMule_evaluateSMSConditions(condObject, index, rowData) {
  var i = index;
  var testHeader = formMule_normalizeHeader(condObject["smsCol-"+i]);
  var testVal = rowData[testHeader];
  var value = condObject["smsVal-"+i];
  if (condObject["smsName-"+i]) {
  var statusCol = formMule_normalizeHeader(condObject["smsName-"+i]+" SMS Status");
    if (rowData[statusCol]!='') {
      var output = false;
      return output;
    }
  }
  var output = false;
  switch(value)
  {
  case "":
      output = true;
      break;
  case "NULL":
      if((!testVal)||(testVal=='')) {
        output = true;
      }  
    break;
  case "NOT NULL":
    if((testVal)&&(testVal!='')) {
        output = true;
      }  
    break;
  default:
    if(testVal==value) {
        output = true;
      }  
  }
  return output;
}


function formMule_evaluateVMConditions(condObject, index, rowData) {
  var i = index;
  var testHeader = formMule_normalizeHeader(condObject["vmCol-"+i]);
  var testVal = rowData[testHeader];
  var value = condObject["vmVal-"+i];
  if (condObject["vmName-"+i]) {
  var statusCol = formMule_normalizeHeader(condObject["vmName-"+i]+" VM Status");
    if (rowData[statusCol]!='') {
      var output = false;
      return output;
    }
  }
  var output = false;
  switch(value)
  {
  case "":
      output = true;
      break;
  case "NULL":
      if((!testVal)||(testVal=='')) {
        output = true;
      }  
    break;
  case "NOT NULL":
    if((testVal)&&(testVal!='')) {
        output = true;
      }  
    break;
  default:
    if(testVal==value) {
        output = true;
      }  
  }
  return output;
}







//returns true if testval meets the condition 
function formMule_evaluateConditions(condObject, index, rowData) {
  var i = index;
  var testHeader = formMule_normalizeHeader(condObject["col-"+i]);
  var testVal = rowData[testHeader];
  var value = condObject["val-"+i];
  if (condObject["sht-"+i]) {
  var statusCol = formMule_normalizeHeader(condObject["sht-"+i]+" Status");
    if (rowData[statusCol]!='') {
      var output = false;
      return output;
    }
  }
  var output = false;
  switch(value)
  {
  case "":
      output = true;
      break;
  case "NULL":
      if((!testVal)||(testVal=='')) {
        output = true;
      }  
    break;
  case "NOT NULL":
    if((testVal)&&(testVal!='')) {
        output = true;
      }  
    break;
  default:
    if(testVal==value) {
        output = true;
      }  
  }
  return output;
}


//returns true if testval meets the condition 
function formMule_evaluateUpdateConditions(condObject, index, rowData) {
  var i = index;
  var testHeader = formMule_normalizeHeader(condObject["col-"+i]);
  var testVal = rowData[testHeader];
  var value = condObject["val-"+i];
  if (condObject["sht-"+i]) {
  var statusCol = formMule_normalizeHeader(condObject["sht-"+i]+" Status");
    if ((rowData[statusCol]!=null)&&(rowData[statusCol]!='')) {
      var output = false;
      return output;
    }
  }
  var output = false;
  switch(value)
  {
  case "":
      output = false;
      break;
  case "NULL":
      if((!testVal)||(testVal=='')) {
        output = true;
      }  
    break;
  case "NOT NULL":
    if((testVal)&&(testVal!='')) {
        output = true;
      }  
    break;
  default:
    if(testVal==value) {
        output = true;
      }  
  }
  return output;
}


function formMule_saveEmailSettings(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.getActiveApplication();
  var emailStatus = e.parameter.emailStatus;
  var emailTrigger = e.parameter.emailTrigger;
  ScriptProperties.setProperty('emailStatus', emailStatus);
  ScriptProperties.setProperty('emailTrigger', emailTrigger);
  var numSelected = e.parameter.numSelectValue;
  var emailConditions = new Object();
  emailConditions["max"]=numSelected;
  for (var i=0; i<numSelected; i++) {
    var condCol = e.parameter["conditionCol-"+i];
    var condVal = e.parameter["value-"+i].trim();
    var templateName = e.parameter["templateName-"+i].trim();
    Logger.log(templateName);
    emailConditions["col-"+i] = condCol;
    emailConditions["val-"+i] = condVal;
    emailConditions["sht-"+i] = templateName;
  }
  emailConditions = Utilities.jsonStringify(emailConditions);
  ScriptProperties.setProperty('emailConditions', emailConditions);
  emailConditions = Utilities.jsonParse(emailConditions);
  ScriptProperties.setProperty('numSelected', numSelected);
  var num = parseInt(numSelected);
  var sheets = ss.getSheets();
  var sheetName = ScriptProperties.getProperty('sheetName');
  var copyDownOption = ScriptProperties.getProperty('copyDownOption');
  var sheet = ss.getSheetByName(sheetName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (var i=0; i<num; i++) {
    var alreadyExists = '';
    for (var j=0; j<sheets.length; j++) {
      if (sheets[j].getName()==emailConditions["sht-"+i]) {
        alreadyExists = true;
        break;
      }
    }
    if (!(alreadyExists==true)) {
    var newSheet = ss.insertSheet().setName(emailConditions["sht-"+i]);
    var newSheetName = newSheet.getName();
    newSheet.getRange(1, 1).setValue("Reply to:").setBackground("yellow");
    newSheet.getRange(2, 1).setValue("To:").setBackground("yellow");
    newSheet.getRange(3, 1).setValue("CC:").setBackground("yellow");
    newSheet.getRange(4, 1).setValue("Subject:").setBackground("yellow"); 
    newSheet.getRange(5, 1).setValue("Body: \n Ctrl+Return gives newline. HTML-friendly!").setVerticalAlignment("top").setBackground("yellow");
    newSheet.getRange(6, 1).setValue("Translate code:").setBackground("yellow").setVerticalAlignment("top");
    newSheet.getRange(5, 2).setVerticalAlignment("top");
    newSheet.setRowHeight(5, 200); 
    newSheet.getRange(1, 3).setValue("<--Optional: Use a single valid email or email token from the list below. Even when this is set, sender address will always appear as the installer of this script.");
    newSheet.getRange(2, 3).setValue("<--Required: Use a single email address or comma separated email addresses, or use email tokens from the list below.");
    newSheet.getRange(3, 3).setValue("<--Optional: Complete me with valid, comma separated email addresses or email tokens from the list below.");
    newSheet.getRange(4, 3).setValue("<--Use tokens below for dynamic values.").setVerticalAlignment("middle");
    newSheet.getRange(5, 3).setValue("<--Use tokens below for dynamic values.").setVerticalAlignment("middle");
    newSheet.getRange(6, 3).setValue("<--Optional: E.g. 'es' for Spanish. Use token for dynamic value. Value must be available in Google Translate: https://developers.google.com/translate/v2/using_rest#language-params");
    newSheet.setColumnWidth(1, 75);
    newSheet.setColumnWidth(2, 350);
    newSheet.setColumnWidth(3, 340);
    formMule_setAvailableTags(newSheet);
  } 
  var headerIndex = headers.indexOf(emailConditions["sht-"+i] + " Status");
  if (headerIndex==-1) {
    var sheet = ss.getSheetByName(sheetName);
    sheet.insertColumnAfter(sheet.getLastColumn());
    sheet.getRange(1, sheet.getLastColumn()+1,1,1).setValue(emailConditions["sht-"+i] + " Status").setFontColor('black').setBackground('#FF9999').setComment("Don't change the name of this column. It is used to log the send status of an email template, whose name must be a match.");
    if (copyDownOption == "true") {
     sheet.getRange(2, sheet.getLastColumn(),1,1).setValue("N/A This is the formula row").setFontColor('black').setBackground('#FF9999'); 
    }
  }    
  }
  app.close();
  return app;
}


function formMule_setAvailableTags(templateSheet) {
  var sheetName = ScriptProperties.getProperty('sheetName');
  var sheet = ss.getSheetByName(sheetName);
  var lastCol = sheet.getLastColumn();
  if(lastCol==0){
    lastCol = 1;
    var noTags = "You have no headers in your data sheet";
  }
  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  var headers = headerRange.getValues();
  var availableTags = [];
  var j=0;
  for (var i=0; i<headers[0].length; i++) {
    if (headers[0][i]=='') {
      Browser.msgBox('Column ' + parseInt(i+1) + ' has a blank header. You cannot have blank headers in your source data. Please fix this before proceeding.');
      break;
    }
    if (headers[0][i]!='Event Update Status') {
      availableTags[j] = [];
      availableTags[j][0] = "${\""+headers[0][i]+"\"}";
      j++; 
    }
  }
  
  var k = availableTags.length;
  availableTags[k] = [];
  availableTags[k][0] = "$currMonth";
  availableTags[k+1] = [];
  availableTags[k+1][0] = "$currDay";
  availableTags[k+2] = [];
  availableTags[k+2][0] = "$currYear";
  availableTags[k+3] = [];
  availableTags[k+3][0] = "$formUrl";
    
  if (noTags) {
    availableTags[0][0] = noTags;
  }
 templateSheet.getRange(7, 1).setValue("Available merge variables").setFontWeight("bold").setBackground("yellow");;
 templateSheet.getRange(7, 2).setValue("Use in any field").setBackground("yellow");
 templateSheet.getRange(8, 1, availableTags.length,1).setBackground("yellow");
 templateSheet.getRange(8, 2, availableTags.length,1).setValues(availableTags).setBackground("yellow");
 templateSheet.getRange(9+availableTags.length, 1).setValue("Handy HTML Tags").setBackground("pink").setFontWeight("bold");
 templateSheet.getRange(9+availableTags.length, 2).setValue("Use in email body only").setBackground("pink");
 var htmlTags = [["Hyperlink","<a href= \"URL\">Link Text</a>"],["Header Level 3","<h3>My header</h3>"],['=hyperlink("http://www.w3schools.com/html/html_basic.asp","Learn more!")','']];
 templateSheet.getRange(10+availableTags.length, 1, htmlTags.length, 2).setValues(htmlTags).setBackground("pink");
 
}



function formMule_sendEmailsAndSetAppointments(manual) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var timeZone = ss.getSpreadsheetTimeZone();
  var properties = ScriptProperties.getProperties();
  var calendarStatus = properties.calendarStatus;
  var calendarUpdateStatus = properties.calendarUpdateStatus;
  var calTrigger = properties.calTrigger;
  var calUpdateTrigger = properties.calUpdateTrigger;
  var copyDownOption = properties.copyDownOption;
  var sheetName = properties.sheetName;
  var sheet = ss.getSheetByName(sheetName);
  ss.setActiveSheet(sheet);
  var headers = formMule_fetchHeaders(sheet);
  if (headers.indexOf("Formula Copy Down Status") != -1) {
    var copyDownFormulasCol = true;
  }
  var days = ["MONDAY", "TUESDAY", "WEDNESDAY", "THURSDAY", "FRIDAY", "SATURDAY", "SUNDAY"];
  if (((calendarStatus == "true")&&(manual==true))||((calendarStatus == "true")&&(calTrigger=="true"))||((calendarUpdateStatus == "true")&&(manual==true))||((calendarUpdateStatus == "true")&&(calUpdateTrigger=="true"))) {
    CacheService.getPrivateCache().remove('lastCalendarId');
    var eventIdCol = properties.eventIdCol;
    var calendarToken  = properties.calendarToken;
    var eventTitleToken = properties.eventTitleToken;
    var locationToken = properties.locationToken;
    var guestsToken = properties.guests;
    var emailInvites = properties.emailInvites;
    if (emailInvites == "true") { emailInvites = true; } else { emailInvites = false;}
    var allDay = properties.allDay;
    var startTimeToken = properties.startTimeToken;
    var endTimeToken = properties.endTimeToken;
    var descToken = properties.descToken;
    var reminderType = properties.reminderType;
    var minBefore = properties.minBefore;
    var calConditionsString = properties.calendarConditions;
    var calCondObject = Utilities.jsonParse(calConditionsString);
    var calUpdateConditionsString = properties.calendarUpdateConditions;
    var calUpdateCondObject = Utilities.jsonParse(calUpdateConditionsString);
    var calDeleteConditionsString = properties.calendarDeleteConditions;
    var calWeeklyTimes = properties.calendarWeeklyRepeats;
    var calRepeats = properties.calendarWeeklyRepeats;
    var calWeekdays = properties.calendarWeekdays;
    var calDeleteCondObject = Utilities.jsonParse(calDeleteConditionsString);
    var eventUpdateStatusIndex = headers.indexOf("Event Update Status");
    var eventCreateStatusIndex = headers.indexOf("Event Creation Status");
    var eventIdIndex = headers.indexOf("Event Id");
    var updateEventIdIndex = headers.indexOf(eventIdCol);
    eventIdCol = formMule_normalizeHeader(eventIdCol);
  }
  var emailStatus = properties.emailStatus;
  var emailTrigger = properties.emailTrigger;
  var emailCondString = properties.emailConditions;
  var smsStatus = properties.smsEnabled;
  var smsTrigger = properties.smsTrigger;
  var vmStatus = properties.vmEnabled;
  var vmTrigger = properties.vmTrigger;
  var accountSID = properties.accountSID;
  var authToken = properties.authToken;
  var maxTexts = properties.smsMaxLength;
      
  if (((emailStatus=="true")&&(manual==true))||((emailStatus=="true")&&(emailTrigger=="true"))) {
    var emailCondString = properties.emailConditions;
    var emailCondObject = Utilities.jsonParse(emailCondString);
    var numSelected = properties.numSelected;
    var templateSheetNames = [];
    var normalizedSendColNames = [];
    var sendersTemplates = [];
    var recipientsTemplates = [];
    var ccRecipientsTemplates = [];
    var subjectTemplates = [];
    var bodyTemplates = [];  
    var langTemplates = [];
    var sendColNames = [];
    for (var i=0; i<numSelected; i++) {
       var templateSheet = ss.getSheetByName(emailCondObject['sht-'+i]);
       sendColNames[i] = emailCondObject['sht-'+i] + " Status";
       normalizedSendColNames[i] = formMule_normalizeHeader(emailCondObject['sht-'+i] + " Status");
       templateSheetNames[i] = templateSheet.getName();
       sendersTemplates[i] = templateSheet.getRange("B1").getValue();
       recipientsTemplates[i] = templateSheet.getRange("B2").getValue();
       ccRecipientsTemplates[i] = templateSheet.getRange("B3").getValue();
       subjectTemplates[i] = templateSheet.getRange("B4").getValue();
       bodyTemplates[i] = templateSheet.getRange("B5").getValue();
       langTemplates[i] = templateSheet.getRange("B6").getValue();
    }
  }
  var lastHeader = sheet.getRange(1, sheet.getLastColumn());
  var lastHeaderValue = formMule_normalizeHeader(lastHeader.getValue());  
  var k=2;
  dataRange = sheet.getRange(k, 1, sheet.getLastRow()-(k-1), sheet.getLastColumn());
  // Create one JavaScript object per row of data.
  var objects = formMule_getRowsData(sheet, dataRange, 1);
  // For every row object, create a personalized email from a template and send
  // it to the appropriate person.
  for (var j = 0; j < objects.length; ++j) {
    if ((properties.copyDownFormulas == "true")&&(!copyDownFormulasCol)) { //self-repair if spreadsheet header for copydown of formulas has been deleted
      returnCopydownStatusColIndex();
      break;
    }
    if (((properties.copyDownFormulas == "true")||(copyDownFormulasCol))&&(objects[j].formulaCopyDownStatus == "")) {
      copyDownFormulas(j+k, properties);
      var thisRowRange = sheet.getRange(j+k, 1, 1, sheet.getLastColumn());
      objects[j] = formMule_getRowsData(sheet, thisRowRange, 1);
    }
    var error = false;
    // Get a row object
    var rowData = objects[j];        
    var confirmation = '';
    var calUpdateConfirmation = '';
    var found = '';
    var calConditionTest = false;
    var calUpdateConditionTest = false;
    var calDeleteConditionTest = false;
    if (calendarStatus=="true") {    
      //test calendar event creation conditions
      if ((calConditionsString)&&(calConditionsString!='')) {
        var calConditionTest = formMule_evaluateConditions(calCondObject, 0, rowData);
      }
    }
    if (calendarUpdateStatus=="true") {
      var calUpdateStatus = rowData.eventUpdateStatus.replace(/ /g,'');
      if (calUpdateStatus=='') {
      //test up calendar update conditions
      if ((calUpdateConditionsString)&&(calUpdateConditionsString!='')) {
        var calUpdateConditionTest = formMule_evaluateUpdateConditions(calUpdateCondObject, 0, rowData);    
      } 
        //test calendar event deletion condition
        if ((calDeleteConditionsString)&&(calDeleteConditionsString!='')) {
          var calDeleteConditionTest = formMule_evaluateUpdateConditions(calDeleteCondObject, 0, rowData);    
        } 
      }
    }
    if ((calendarToken)&&((calUpdateConditionTest)||(calConditionTest)||(calDeleteConditionTest))) {
      var lastCalendarId = CacheService.getPrivateCache().get('lastCalendarId');
      var calendarId  = formMule_fillInTemplateFromObject(calendarToken, rowData);
      if (lastCalendarId!=calendarId) {
        var calendar = CalendarApp.getCalendarById(calendarId);
      }
    }
     
    //Either the event is being created or it's being updated
    if (((calConditionTest == true)||(calUpdateConditionTest == true))&&(calendar)) {
        var eventTitle = formMule_fillInTemplateFromObject(eventTitleToken, rowData);
        var location = formMule_fillInTemplateFromObject(locationToken, rowData);
        var desc = formMule_fillInTemplateFromObject(descToken, rowData);
        var guests = formMule_fillInTemplateFromObject(guestsToken, rowData);
        var repeats = formMule_fillInTemplateFromObject(calRepeats, rowData);
        var weekdays = formMule_fillInTemplateFromObject(calWeekdays, rowData);
        var startTimeStamp;
        var endTimeStamp;
        if (!(allDay=="true")) {
          var startTimeStamp = new Date(formMule_fillInTemplateFromObject(startTimeToken, rowData));
          var endTimeStamp = new Date(formMule_fillInTemplateFromObject(endTimeToken, rowData));
        } else {
          startTimeStamp = new Date(formMule_fillInTemplateFromObject(startTimeToken, rowData));
        }
        if (guests!='') {
          var options = {guests:guests, location:location, description:desc, sendInvites:emailInvites}; 
        } else {
          var options = {location:location, description:desc,};
        }
        //calendar event creation condition
        if ((calConditionTest == true)&&(calUpdateConditionTest == false)) {
          var event = '';
          var eventId = ''; 
          var calConfirmation = '';     
        try {
          if ((allDay=="true")&&(startTimeStamp != "Invalid Date")) {
            if ((!repeats)||(repeats=='')||(repeats==1)) {
              eventId = calendar.createAllDayEvent(eventTitle, startTimeStamp, options).getId();
              calConfirmation += "All day event added to calendar: "+ calendar.getName() + " for " + startTimeStamp;
            } else {
              if (repeats) {
                repeats = parseInt(repeats); 
                if (weekdays) {
                  weekdays = weekdays.split(",");
                  var weekdayArray = []
                  for (var h=0; h<weekdays.length; h++) {
                    var weekday = weekdays[h].trim();
                    weekday = weekday.toUpperCase();
                    if (days.indexOf(weekday)!=-1) {
                      weekdayArray.push(CalendarApp.Weekday[weekday]);
                    }
                  }
                  if (weekdayArray.length>0) { 
                    repeats = weekdayArray.length * repeats;
                    eventId = calendar.createAllDayEventSeries(eventTitle, startTimeStamp, CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekdays(weekdayArray).times(repeats), options).getId();
                  }
                } else {
                 eventId = calendar.createAllDayEventSeries(eventTitle, startTimeStamp, CalendarApp.newRecurrence().addDailyRule().interval(7).times(repeats), options).getId();
                }
              }
              calConfirmation += "Recurring all day event added to calendar: "+ calendar.getName() + " starting " + startTimeStamp;
            }
            try {
              formMule_logCalEvent();
            } catch(err) {
            }
            sheet.getRange(j+k, eventIdIndex+1).setValue(eventId);
          } else if (startTimeStamp != "Invalid Date") {
             if ((!repeats)||(repeats=='')||(repeats==1)) {
               eventId = calendar.createEvent(eventTitle, startTimeStamp, endTimeStamp, options).getId();
               calConfirmation += "Event added to calendar: "+ calendar.getName() + " for " + startTimeStamp;
             } else {
                if (repeats) {
                  repeats = parseInt(repeats); 
                    if (weekdays) {
                      weekdays = weekdays.split(",");
                      var weekdayArray = []
                      for (var h=0; h<weekdays.length; h++) {
                        var weekday = weekdays[h].trim();
                        weekday = weekday.toUpperCase();
                        if (days.indexOf(weekday)!=-1) {
                          weekdayArray.push(CalendarApp.Weekday[weekday]);
                        }
                      }
                      if (weekdayArray.length>0) { 
                        repeats = weekdayArray.length * repeats;
                        eventId = calendar.createEventSeries(eventTitle, startTimeStamp, endTimeStamp, CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekdays(weekdayArray).times(repeats), options).getId();
                      }
                    } else {
                     eventId = calendar.createEventSeries(eventTitle, startTimeStamp, endTimeStamp, CalendarApp.newRecurrence().addDailyRule().interval(7).times(repeats), options).getId();
                    }
                  }
               calConfirmation += "Recurring event added to calendar: "+ calendar.getName() + " starting " + startTimeStamp;
            }
            try {
              formMule_logCalEvent();
            } catch(err) {
            }
            sheet.getRange(j+k, eventIdIndex+1).setValue(eventId);
          } else {
            calConfirmation = "No event created. Invalid date/time parameters.";
            error = true;
          }
          } catch(err) {
            calConfirmation = "Create calendar event failed: " + err;
            error = true;
          }
        if (eventId) {
          event = calendar.getEventSeriesById(eventId);
          if (reminderType) { 
            switch(reminderType) {
            case "Email reminder": 
              event = calendar.getEventSeriesById(eventId).addEmailReminder(minBefore);
            break;
            case "Popup reminder":
              event = calendar.getEventSeriesById(eventId).addPopupReminder(minBefore);
            break;
            case "SMS reminder":
              event = calendar.getEventSeriesById(eventId).addSmsReminder(minBefore);
            break;
            default:
              event = calendar.getEventSeriesById(eventId).removeAllReminders();
            }
         }
         rowData.eventId = eventId;
         desc = formMule_fillInTemplateFromObject(descToken, rowData);
         event = calendar.getEventSeriesById(eventId).setDescription(desc); 
       }
      sheet.getRange(j+k, eventCreateStatusIndex+1).setValue(calConfirmation).setFontColor("black");
      if (error==true) {
        sheet.getRange(j+k, eventCreateStatusIndex+1).setFontColor("red");
      }
      error = false;
     } 
    }
    if ((calUpdateConditionTest == true)||(calDeleteConditionTest == true)) {
      var eventId = rowData[eventIdCol];
      found = false;
      try {
        var future = new Date();
        future.setDate(future.getDate()+360);
        var past = new Date();
        past.setDate(past.getDate()-360);
        var events = calendar.getEvents(past, future);
        for (var g = 0; g<events.length; g++) {
          if (events[g].getId()==eventId) {
            var event = events[g];
            found = true;
          }
        }
      } catch(err) {
        found = false;
      }
    }
    if ((calUpdateConditionTest == true)&&(found == true)) {
      try {
        if (allDay!="true") {
          if ((!repeats)||(repeats=='')||(repeats==1)) {
            var event = event.setTime(startTimeStamp, endTimeStamp);
          } else {
              if (repeats) {
                  repeats = parseInt(repeats); 
                    if (weekdays) {
                      weekdays = weekdays.split(",");
                      var weekdayArray = []
                      for (var h=0; h<weekdays.length; h++) {
                        var weekday = weekdays[h].trim();
                        weekday = weekday.toUpperCase();
                        if (days.indexOf(weekday)!=-1) {
                          weekdayArray.push(CalendarApp.Weekday[weekday]);
                        }
                      }
                      if (weekdayArray.length>0) { 
                        repeats = weekdayArray.length * repeats;
                        event = calendar.getEventSeriesById(eventId).setRecurrence(CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekdays(weekdayArray).times(repeats), startTimeStamp, endTimeStamp);
                      }
                    } else {
                      event = calendar.getEventSeriesById(eventId).setRecurrence(CalendarApp.newRecurrence().addDailyRule().interval(7).times(repeats), startTimeStamp, endTimeStamp);
                    }
                  }
          }
        } else {
          if ((!repeats)||(repeats=='')||(repeats==1)) {
            var event = event.setAllDayDate(startTimeStamp);
          } else {
              if (repeats) {
                  repeats = parseInt(repeats); 
                    if (weekdays) {
                      weekdays = weekdays.split(",");
                      var weekdayArray = []
                      for (var h=0; h<weekdays.length; h++) {
                        var weekday = weekdays[h].trim();
                        weekday = weekday.toUpperCase();
                        if (days.indexOf(weekday)!=-1) {
                          weekdayArray.push(CalendarApp.Weekday[weekday]);
                        }
                      }
                      if (weekdayArray.length>0) { 
                        repeats = weekdayArray.length * repeats;
                        event = calendar.getEventSeriesById(eventId).setRecurrence(CalendarApp.newRecurrence().addWeeklyRule().onlyOnWeekdays(weekdayArray).times(repeats), startTimeStamp);
                      }
                    } else {
                      event = calendar.getEventSeriesById(eventId).setRecurrence(CalendarApp.newRecurrence().addDailyRule().interval(7).times(repeats), startTimeStamp);
                    }
                  }
          }
        }
        event = calendar.getEventSeriesById(eventId).setDescription(desc)
        event = calendar.getEventSeriesById(eventId).setLocation(location);
        event = calendar.getEventSeriesById(eventId).setTitle(eventTitle);
        event = calendar.getEventSeriesById(eventId).removeAllReminders();
        calUpdateConfirmation += eventTitle + ' successfully moved/updated to ' + startTimeStamp;
        try {
          formMule_logCalUpdate();
        } catch(err) {
        }
      } catch(err) {
        calUpdateConfirmation += eventId + ' not found on calendar: ' + calendar.getName();
      }
      sheet.getRange(j+k, eventUpdateStatusIndex+1).setValue(calUpdateConfirmation).setFontColor("black");
      if (found==false) {
        sheet.getRange(j+k, eventUpdateStatusIndex+1).setFontColor("red");
      }
      if ((eventId)&&(eventId!="")&&(found==true)) {
        event = calendar.getEventSeriesById(eventId);    
        if (reminderType) { 
        switch(reminderType) {
        case "Email reminder": 
          event = calendar.getEventSeriesById(eventId).addEmailReminder(minBefore);
        break;
        case "Popup reminder":
          event = calendar.getEventSeriesById(eventId).addPopupReminder(minBefore);
        break;
        case "SMS reminder":
          event = calendar.getEventSeriesById(eventId).addSmsReminder(minBefore);
        break;
        default:
          event = calendar.getEventSeriesById(eventId).removeAllReminders();
        }
       }
       rowData.eventId = eventId;
       desc = formMule_fillInTemplateFromObject(descToken, rowData);
       event = calendar.getEventSeriesById(eventId).setDescription(desc); 
     }
     if (guests!='') {
       guests = guests.split(",");
       for (var g=0; g<guests.length; g++) {
         event.addGuest(guests[g]);
       }
      }
    }
  if ((found==true)&&(calDeleteConditionTest == true)&&(calUpdateConditionTest == false)) {
    var eventTitle = event.getTitle();
    event = calendar.getEventSeriesById(eventId).deleteEventSeries();
    calUpdateConfirmation = eventTitle + ' deleted from calendar: ' + calendar.getName();
    sheet.getRange(j+k, eventUpdateStatusIndex+1).setValue(calUpdateConfirmation).setFontColor("black");
  }
  if ((found==false)&&(calendar)&&((calUpdateConditionTest == true)||(calDeleteConditionTest == true))) {
     calUpdateConfirmation = eventId + ' not found on calendar: ' + calendar.getName();
     sheet.getRange(j+k, eventUpdateStatusIndex+1).setValue(calUpdateConfirmation);
     sheet.getRange(j+k, eventUpdateStatusIndex+1).setFontColor("red");
  }
  if ((!(calendar))&&(calUpdateConditionTest == true)) {
     calUpdateConfirmation = "Calendar not found";
     sheet.getRange(j+k, eventUpdateStatusIndex+1).setValue(calUpdateConfirmation);
     sheet.getRange(j+k, eventUpdateStatusIndex+1).setFontColor("red");
  }



  if (((emailStatus=="true")&&(manual==true))||((emailStatus=="true")&&(emailTrigger=="true"))) {
  for (var i=0; i<numSelected; i++) {
    var confirmation = '';
    if ((emailCondString)&&(emailCondString!='')) {
      var emailConditionTest = formMule_evaluateConditions(emailCondObject, i, rowData);
    }
    if ((emailConditionTest == true)||(!emailCondString)) {
      var sendersTemplate = sendersTemplates[i];
      var recipientsTemplate = recipientsTemplates[i];
      var ccRecipientsTemplate = ccRecipientsTemplates[i];
      var subjectTemplate = subjectTemplates[i];
      var bodyTemplate = bodyTemplates[i];
      var langTemplate = langTemplates[i];

    // Generate a personalized email.
    // Given a template string, replace markers (for instance ${"First Name"}) with
    // the corresponding value in a row object (for instance rowData.firstName).
    var from = formMule_fillInTemplateFromObject(sendersTemplate, rowData);
    var to = formMule_fillInTemplateFromObject(recipientsTemplate, rowData);
    try {
      if (to!='') {
        var cc = formMule_fillInTemplateFromObject(ccRecipientsTemplate, rowData);
        var subject = formMule_fillInTemplateFromObject(subjectTemplate, rowData);
        var body = formMule_fillInTemplateFromObject(bodyTemplate, rowData); 
        var lang = formMule_fillInTemplateFromObject(langTemplate, rowData);
        body = body.replace("\n", "<br />","g");
        if ((lang)&&(lang!='')) {
          var translation = LanguageApp.translate("Automated Translation", '', lang);
          var divider = '<h3>####### ' + translation + ' #######</h3>';
          body += divider + LanguageApp.translate(body, '', lang)
        }
      if (from=='') {
        MailApp.sendEmail(to, subject, '', {htmlBody: body, cc: cc});
      } else {
        MailApp.sendEmail(to, subject, '', {htmlBody: body, cc: cc, replyTo: from});
      } 
      var now = new Date();
      now = Utilities.formatDate(now, timeZone, "MM/dd/yy' at 'h:mm:ss a")
      if (cc!='') { var ccMsg = ", and cc'd to "+cc; } else {ccMsg = ''}
      if ((i>0)&&(i<numSelected-1)) { var addSemiColon = "; "} else { var addSemiColon = ""; }
        confirmation += templateSheetNames[i]+" sent to "+to+ccMsg+' on '+ now + addSemiColon;
        try {
          formMule_logEmail();
        } catch(err) {
        }
      } else { 
        confirmation += templateSheetNames[i] + " error: Template missing \"To\" address";
        var error = true;
      }
    } catch(err) {
      confirmation += templateSheetNames[i] +" error: " + err;
      var error = true;
    }
    var statusIndex = headers.indexOf(sendColNames[i]);
    var statusCell = sheet.getRange(j+k, statusIndex+1);
    if (confirmation!='') {
    var statusCell = sheet.getRange(j+k, statusIndex+1);
      statusCell.setValue(confirmation).setFontColor("black");
    }
    if (error==true) {
      statusCell.setFontColor("red");
    }
    error == false
   } // end per email conditional check
  } // end i loop through email templates
  } // end conditional test for confirmation email
    //begin SMS section
    if (((smsStatus=="true")&&(manual==true))||((smsStatus=="true")&&(smsTrigger=="true"))) { 
      var twilioNumber = properties.twilioNumber;
      var smsPropertyString = properties.smsPropertyString;
      for (var i=0; i<properties.smsNumSelected; i++) {
        if ((smsPropertyString)&&(emailCondString!='')) {
          var smsPropertyObject = Utilities.jsonParse(smsPropertyString);
          var smsConditionTest = formMule_evaluateSMSConditions(smsPropertyObject, i, rowData);  
          if ((smsConditionTest == true)||(!smsPropertyString)) {
            var phoneNumber = formMule_fillInTemplateFromObject(smsPropertyObject['smsPhone-'+i], rowData);
            var lang = formMule_fillInTemplateFromObject(smsPropertyObject['smsLang-'+i], rowData);
            lang = lang.trim();
            var body = formMule_fillInTemplateFromObject(smsPropertyObject['smsBody-'+i], rowData, true);
            var flagObject = new Object();
            flagObject.sheetName = sheetName;
            flagObject.header = urlencode(smsPropertyObject['smsName-'+i] + " SMS Status");
            flagObject.row = j+k;
            var args = new Object();
            formMule_sendAText(phoneNumber, body, accountSID, authToken, flagObject, args, lang, maxTexts);
            try {
              formMule_logSMS();
            } catch(err) {
            }
          }
        }
      }
    }
    //end SMS Section
    //begin VM section
    if (((vmStatus=="true")&&(manual==true))||((vmStatus=="true")&&(vmTrigger=="true"))) {
      var vmPropertyString = properties.vmPropertyString;
      for (var i=0; i<properties.vmNumSelected; i++) {
        if ((vmPropertyString)&&(emailCondString!='')) {
          var vmPropertyObject = Utilities.jsonParse(vmPropertyString);
          var vmConditionTest = formMule_evaluateVMConditions(vmPropertyObject, i, rowData);  
          if ((vmConditionTest == true)||(!vmPropertyString)) {
            var phoneNumber = formMule_fillInTemplateFromObject(vmPropertyObject['vmPhone-'+i], rowData);
            var lang = formMule_fillInTemplateFromObject(vmPropertyObject['vmLang-'+i], rowData);
            lang = lang.trim();
            var body = formMule_fillInTemplateFromObject(vmPropertyObject['vmBody-'+i], rowData, true);
            var flagObject = new Object();
            flagObject.sheetName = sheetName;
            flagObject.header = urlencode(vmPropertyObject['vmName-'+i] + " VM Status");
            flagObject.row = j+k;
            var args = new Object();
            if (vmPropertyObject['vmRecordOption-'+i]=="Yes") {
              args.RequestResponse = "TRUE";
            }
            if ((vmPropertyObject['vmSoundFile-'+i])&&(vmPropertyObject['vmSoundFile-'+i])!='') {  
              args.PlayFile = vmPropertyObject['vmSoundFile-'+i];
            }
            formMule_makeRoboCall(phoneNumber, body, accountSID, authToken, flagObject, args, lang);
            try {
              formMule_logVoiceCall();
            } catch(err) {
            }
          }
        }
      }
    } 
    //send VM section
  } //end j loop through spreadsheet rows
} // end function



// Replaces markers in a template string with values define in a JavaScript data object.
// Arguments:
//   - template: string containing markers, for instance ${"Column name"}
//   - data: JavaScript object with values to that will replace markers. For instance
//           data.columnName will replace marker ${"Column name"}
// Returns a string without markers. If no data is found to replace a marker, it is
// simply removed.
function formMule_fillInTemplateFromObject(template, data, newline) {
  var newString = template;
  // Search for all the variables to be replaced, for instance ${"Column name"}
  var templateVars = template.match(/\$\{\"[^\"]+\"\}/g);
  // Replace variables from the template with the actual values from the data object.
  // If no value is available, replace with the empty string.
  if (templateVars) {
  for (var i = 0; i < templateVars.length; ++i) {
    // formMule_normalizeHeader ignores ${"} so we can call it directly here.
    var variableData = data[formMule_normalizeHeader(templateVars[i])];
    newString = newString.replace(templateVars[i], variableData || "");
  }
  }
  var currentTime = new Date();
  var month = (currentTime.getMonth() + 1).toString();
  var day = currentTime.getDate().toString();
  var year = currentTime.getFullYear().toString();
  
  if (!newline) {  newString = newString.replace("\n", "<br />","g"); }
  if (newString.indexOf("$currMonth")!=-1) {
    newString = newString.replace("$currMonth", month);
  }
  if (newString.indexOf("$currDay")!=-1) {
    newString = newString.replace("$currDay", day);
  }
  if (newString.indexOf("$currYear")!=-1) {
    newString = newString.replace("$currYear", year);
  }
  if (newString.indexOf("$formUrl")!=-1) {
    var formUrl = CacheService.getPrivateCache().get('formUrl');
      if (!formUrl) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      formUrl = ScriptProperties.getProperty('formUrl');
      CacheService.getPrivateCache().put('formUrl', formUrl);
    }
    newString = newString.replace("$formUrl", formUrl);
  }
  if (data.eventId) {
    newString = newString.replace("$eventId", data.eventId);
  }
  return newString;
}





//////////////////////////////////////////////////////////////////////////////////////////
//
// The code below is reused from the 'Reading Spreadsheet data using JavaScript Objects'
// tutorial.
//
//////////////////////////////////////////////////////////////////////////////////////////

// formMule_getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range; 
// Returns an Array of objects.
function formMule_getRowsData(sheet, range, columnHeadersRowIndex) {
  columnHeadersRowIndex = columnHeadersRowIndex || range.getRowIndex() - 1;
  var numColumns = range.getEndColumn() - range.getColumn() + 1;
  var headersRange = sheet.getRange(columnHeadersRowIndex, range.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return formMule_getObjects(range.getValues(), formMule_normalizeHeaders(headers));
}

// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function formMule_getObjects(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
    //  if (formMule_isCellEmpty(cellData)) {
    //    continue;
    //  }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Arguments:
//   - headers: Array of Strings to normalize
function formMule_normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    var key = formMule_normalizeHeader(headers[i]);
    if (key.length > 0) {
      keys.push(key);
    }
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function formMule_normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!formMule_isAlnum(letter)) {
      continue;
    }
    if (key.length == 0 && formMule_isDigit(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function formMule_isCellEmpty(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function formMule_isAlnum(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    formMule_isDigit(char);
}

// Returns true if the character char is a digit, false otherwise.
function formMule_isDigit(char) {
  return char >= '0' && char <= '9';
}



function formMule_defineSettings() {
  setSid();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var app = UiApp.createApplication().setTitle("Step 1: Define merge source settings").setHeight(300);
  var panel = app.createVerticalPanel().setId("settingsPanel");  
  var sheetLabel = app.createLabel().setText("Choose the sheet that you want the script to use as a data source for merged emails and/or calendar events.");
  sheetLabel.setStyleAttribute("background", "#E5E5E5").setStyleAttribute("marginTop", "20px").setStyleAttribute("padding", "5px");
  var sheetChooser = app.createListBox().setId("sheetChooser").setName("sheet");
    for (var i=0; i<sheets.length; i++) {
      if ((!sheets[i].getName().match("Template"))&&(!sheets[i].getName().match("formMule Read Me"))&&(!sheets[i].getName().match("Forms in same folder"))) {
      sheetChooser.addItem(sheets[i].getSheetName());   
      }
    }
   if (ScriptProperties.getProperty('sheetName')) {
   var sheetName = ScriptProperties.getProperty('sheetName');
   }
  if (sheetName) {
    var sheetIndex = formMule_getSheetIndex(sheetName);
    sheetChooser.setSelectedIndex(sheetIndex);
  }
  
  var optionsLabel = app.createLabel("Optional").setStyleAttribute("background", "#E5E5E5").setStyleAttribute("marginTop", "20px").setStyleAttribute("padding", "5px");
 
  var caseNoSetting = ScriptProperties.getProperty('caseNoSetting'); 
  var caseNoCheckBox = app.createCheckBox().setId("caseNoCheckBox").setName("caseNoSetting").setText("Auto-create a unique case number for each form submission")
  if (caseNoSetting=="true") {
   caseNoCheckBox.setValue(true);
  }
  var helpPopup = app.createPopupPanel();
  var fieldsLabel = app.createLabel().setText("The formMule script can be used to merge emails and calendar events from any data sheet, however it gains additional power when coupled with a Google Form, VLOOKUP, and other formulas.  See the \"Read Me\" tab in this spreadsheet or visit http://www.youpd.org/formmule for more information about how to use it.");                                         

                                          
  helpPopup.add(fieldsLabel);
  panel.add(helpPopup);
  panel.add(sheetLabel);
  panel.add(sheetChooser);
  panel.add(optionsLabel);
  panel.add(caseNoCheckBox);


  var spinner = app.createImage(MULEICONURL).setWidth(150);
  spinner.setVisible(false);
  spinner.setStyleAttribute("position", "absolute");
  spinner.setStyleAttribute("top", "120px");
  spinner.setStyleAttribute("left", "190px");
  spinner.setId("dialogspinner");
  
  var copyDownLabel = app.createLabel("Looking to copy formulas down on form submit? This feature is now located in the \"Advanced options\" menu option.").setStyleAttribute('marginTop', '10px');
  panel.add(copyDownLabel);
  var spinnerHandler = app.createClientHandler().forTargets(spinner).setVisible(true).forTargets(panel).setStyleAttribute('opacity', '0.5');
  app.add(panel);
  
  var buttonHandler = app.createServerClickHandler('formMule_saveSettings').addCallbackElement(panel);
  var button = app.createButton("Save settings", buttonHandler).setId('saveButton');
  button.addMouseDownHandler(spinnerHandler); 
  panel.add(button);
  app.add(spinner);
  this.ss.show(app);
}


function formMule_saveSettings(e) {
  var app = UiApp.getActiveApplication();
  var oldSheetName = ScriptProperties.getProperty('sheetName');
  var sheetName = e.parameter.sheet;
  ScriptProperties.setProperty('sheetName', sheetName);
  var sheet = ss.getSheetByName(sheetName);
  var headers = formMule_fetchHeaders(sheet);
  var lastCol = sheet.getLastColumn();
  
  if (lastCol==0) { 
    Browser.msgBox("You have no headers in your data sheet. Please fix this and come back."); 
    formMule_defineSettings();
    app.close();
    return app; 
  }
    
  var caseNoSetting = e.parameter.caseNoSetting;
  ScriptProperties.setProperty('caseNoSetting', caseNoSetting);
  
  var caseNo = ScriptProperties.getProperty('caseNo');
  var caseNoIndex = headers.indexOf("Case No");
  
  if (caseNoSetting=="true") {
  if (caseNoIndex == -1) {
    sheet.insertColumnAfter(sheet.getLastColumn());
    sheet.getRange(1,sheet.getLastColumn()+1).setBackground("orange").setFontColor("black").setValue("Case No").setComment("Don't change the name of this column. It is used to log a unique case number for each form submission.");
   }
  }
  if (caseNoSetting=="false") {
    if ((caseNoIndex != -1)) {
    sheet.deleteColumn(caseNoIndex+1);
    }
  }
  
  var lastCol = sheet.getLastColumn();
  
  var sheetName = ScriptProperties.getProperty('sheetName');
  if ((sheetName)&&(!(oldSheetName))) {
    onOpen();
    Browser.msgBox('Auto-email and auto-calendar-event options are now available in the formMule menu');
  }
  onOpen();
  app.close()
  return app;
}



function formMule_getSheetIndex(sheetName) {
  var app = UiApp.getActiveApplication();
  var sheets = ss.getSheets();
  var bump = 0;
  for (var i=0; i<sheets.length; i++) {
    if (sheets[i].getName()==sheetName) {
     var index = i;
     break;
    }
     if ((sheets[i].getName()== "formMule Read Me")||(sheets[i].getName()=="Email1 Template")||(sheets[i].getName()== "Email2 Template")||(sheets[i].getName()=="Email3 Template")||(sheets[i].getName()== "Forms in same folder"))  {
        bump = bump-1;
     }
    index = 0;
  }
index = index+bump;
return index;
}

function formMule_refreshGrid (e) {
  var app = UiApp.getActiveApplication();
  var panel = app.getElementById('settingsPanel');
  var gridPanel = app.getElementById('gridPanel');
  gridPanel.setStyleAttribute('opacity', '1');
  var grid = app.getElementById("checkBoxGrid");
  var spinner = app.getElementById("dialogspinner");
  spinner.setVisible(false);
  panel.setStyleAttribute('opacity', '1');
  var sheetName = e.parameter.sheet;
  var sheet = ss.getSheetByName(sheetName);
  var headers = formMule_fetchHeaders(sheet);
  grid.resize(headers.length, 1);
  var copyDownOption = e.parameter.copyDownOption;
  if (copyDownOption=="true") {
    gridPanel.setVisible(true);
    gridPanel.setEnabled(true);
  } else {
    gridPanel.setVisible(false);
    gridPanel.setEnabled(false);
  }
  var count = 0;
  for (var i=0; i<headers.length; i++) {
    var normalizedHeader = formMule_normalizeHeader(headers[i]);
    var checkBox = app.createCheckBox().setText(headers[i]).setName(normalizedHeader).setId("checkBox-"+i);
    var cell = sheet.getRange(1, i+1);
    if ((cell.getBackground()=="#DDDDDD")||(cell.getBackground()=="black")||(cell.getBackground()=="orange")||(cell.getBackground()=="#0099FF")||(cell.getBackground()=="#FF9999")||(cell.getBackground()=="#66C285")) { 
      checkBox.setVisible(false); 
    } else {
      count++;
    }
    grid.setWidget(i,0, checkBox); 
   }
  if (count==0) {
    var noneLabel = app.createLabel("You have no columns to the right of your form data.");
    grid.setWidget(0,0, noneLabel);
  }
  return app;
}
var excludedHeaders = ['Merged Doc ID','Merged Doc URL','Link to merged Doc','Document Merge Status',"Case No","Formula Copy Down Status"];


function returnAlphabet() {
  var alphabet = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"];
  var properties = ScriptProperties.getProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getSheetById(ss, properties.formSheetId);
  var excelAlphabet = [];
  for (var i=0; i<sheet.getLastColumn(); i++) {
    var thisLetter = '';
    var numIterations = i/26;
    var mod = i%26;
    var thisEndLetter = alphabet[mod];
    var thisLeadingLetter = '';
    if (numIterations>=1) {
      thisLeadingLetter = alphabet[Math.floor(numIterations)];
    }
    thisLetter = thisLeadingLetter + thisEndLetter;
    excelAlphabet.push(thisLetter);
  }
  return excelAlphabet;
}



function getAllExludedHeaders() {
  var theseExcludedHeaders = excludedHeaders;
  var emailConditions = ScriptProperties.getProperty('emailConditions');
  if (emailConditions) {
    emailConditions = Utilities.jsonParse(emailConditions);
    for (var key in emailConditions) {
      if (key.split("-")[0] == "sht") {
        theseExcludedHeaders.push(emailConditions[key] + " Status");
      }
    }
  }
  var calendarConditions = ScriptProperties.getProperty('calendarConditions');
  if (calendarConditions) {
    calendarConditions = Utilities.jsonParse(calendarConditions);
    for (var key in calendarConditions) {
      if (key.split("-")[0] == "sht") {
        theseExcludedHeaders.push(calendarConditions[key] + " Status");
      }
    }
  }
  var smsConditions = ScriptProperties.getProperty('smsConditions');
  if (smsConditions) {
    smsConditions = Utilities.jsonParse(smsConditions);
    for (var key in smsConditions) {
      if (key.split("-")[0] == "sht") {
        theseExcludedHeaders.push(smsConditions[key] + " Status");
      }
    }
  }
  var formUrl = ScriptProperties.getProperty('formUrl')
  var form = FormApp.openByUrl(formUrl);
  var items = form.getItems();
  var excludedTypes = ['PAGE_BREAK','SECTION_HEADER','IMAGE']
  for (var i=0; i<items.length; i++) {
    var type = items[i].getType().toString();
    if (excludedTypes.indexOf(type) == -1) {
      theseExcludedHeaders.push(items[i].getTitle());
    }
  }
  return theseExcludedHeaders;
}


function copyDownFormulas(thisRow, properties) {
  if ((properties.copyDownFormulas == "true")&&(thisRow>1)) {
    var alphabet = returnAlphabet()
    var statusCol = returnCopydownStatusColIndex()+1;
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getSheetById(ss, properties.formSheetId);
    var colValues = Utilities.jsonParse(properties.colValues);
    var formulaRow = parseInt(properties.formulaRow);
    var objArray = [];
    for (var key in colValues) {
      var colNum = key.split("-")[1];
      objArray.push({colNum: colNum, type: colValues[key], formula: properties['formula-'+colNum]});
    }
    objArray.sort(function(a,b){return a.colNum-b.colNum});
    var colNums = '';
    for (var i=0; i<objArray.length; i++) {
      var colNum = objArray[i].colNum;
      if (i>0) {
        colNums += ",";
      }
      colNums += alphabet[objArray[i].colNum-1];
      var type = objArray[i].type;
      var formula = objArray[i].formula;
      var copyCell = sheet.getRange(formulaRow, colNum).setFormula(formula);
      var destCell = sheet.getRange(thisRow, colNum);
      copyCell.copyTo(destCell);
      if (type == 1) {
        var newValue = destCell.getValue();
        var tries = 0;
        while ((newValue == "#N/A")&&(tries<5)) {
          SpreadsheetApp.flush();
          newValue = destCell.getValue();
          Utilities.sleep(tries * 1000);
          tries++;
        } 
        destCell.setValue(newValue);
      }
    }
    var now = new Date();
    now = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "MM/dd/yy' at 'h:mm:ss a")
    sheet.getRange(thisRow, statusCol).setValue("Formulas in columns " + colNums + " copied down on " + now);
    SpreadsheetApp.flush();
  }
  return;
}



function returnCopydownStatusColIndex() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = ScriptProperties.getProperty('sheetName');
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var index = headers.indexOf("Formula Copy Down Status");
    if (index == -1) {
      var lastCol = sheet.getLastColumn();
      sheet.insertColumnAfter(lastCol);
      index = lastCol + 1;
      sheet.getRange(1, index).setValue("Formula Copy Down Status").setNote("Edit formula copy down preferences in Advanced settings").setBackground("purple").setFontColor("white").setFontStyle("bold");
    }
    return index;
  }
  return -1;
}



function formMule_detectFormSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();
  ScriptProperties.setProperty('ssId', ssId);
  var formUrl = ScriptProperties.getProperty('formUrl');
  if (!formUrl) {
    Browser.msgBox("This feature only works on Spreadsheets with attached forms");
    return;
  }
  try {
    var form = FormApp.openByUrl(formUrl);
  } catch(err) {
    Browser.msgBox("Unable to access the form attached to this spreadsheet...");
  }
  var items = form.getItems();
  var sheets = ss.getSheets();
  for (var i=0; i<sheets.length; i++) {
    var thisTopLeftCell = sheets[i].getRange(1, 1);
    var thisTopLeftBgColor = thisTopLeftCell.getBackground();
    var thisTopLeftValue = thisTopLeftCell.getValue();
    var found = false;
    if ((thisTopLeftBgColor == "#DDDDDD")&&(thisTopLeftValue == "Timestamp")) {
      var formSheetId = sheets[i].getSheetId();
      found = true;
      break;
    }
  }
  if (found) {
    ScriptProperties.setProperty('formSheetId', formSheetId);
  } else {
    var error = catchNoFormSheetDetected(ss);
  }
  if ((!error)||(error!=false)) {
    formulaCaddy_createJob();
  }
}


function catchNoFormSheetDetected(ss) {
  var formSheetName = Browser.inputBox('Unable to detect form sheet, please enter the name of the sheet that holds your form responses');
  try {
    var formSheetId = ss.getSheetByName(formSheetName).getSheetId();
  } catch(err) {
    if (formSheetName!="cancel") {
      Browser.msgBox('Unable to find sheet: ' + formSheetName);
      catchNoFormSheetDetected(ss);
    } else {
      return 'error';
    }
  }
  ScriptProperties.setProperty('formSheetId', formSheetId);
  return;
}


function getSheetById(ss, sheetId) {
  var sheets = ss.getSheets();
  for (var i=0; i<sheets.length; i++) {
    if (sheets[i].getSheetId() == sheetId) {
      return sheets[i];
    }
  }
  return;
}


function formulaCaddy_createJob() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  var sheets = ss.getSheets();
  var sheetNames = [];
  for (var i=0; i<sheets.length; i++) {
    sheetNames.push(sheets[i].getName());
  }
  var formSheet = getSheetById(ss, properties.formSheetId);
  var activeSheetIndex = formSheet.getIndex();
  var activeSheetName = formSheet.getName();
  var activeSheetSelectIndex = sheetNames.indexOf(activeSheetName);
  var app = UiApp.createApplication().setTitle("Copy down formulas within form response sheet").setHeight(400).setWidth(700);
  var outerScroll = app.createScrollPanel().setHeight("390px").setWidth("690px");
  var panel = app.createVerticalPanel().setId("panel").setWidth("680px").setHeight("200px");
  var help = "This feature allows you to use additional columns in the form response sheet to calculate or look up values upon form submission. ";
  help += "Only columns that are not part of your form will be available for selection below. Selecting the \"Paste as values?\" checkbox will ";
  help += "paste the calculated values into the form submission row, without formulas."  
  help += "<br/> Be aware that any changes you make to the column order or structure of the form responses sheet will require you to redo these settings."; 
  var helpHtml = app.createHTML(help).setStyleAttribute('marginBottom', '5px');
  panel.add(helpHtml);
  var columns = formSheet.getLastColumn();
  var grid = app.createGrid(4, columns+1).setId('grid');
  var spinner = app.createImage(MULEICONURL).setWidth("115px").setId('spinner');
  spinner.setVisible(false);
  spinner.setStyleAttribute("position", "absolute");
  spinner.setStyleAttribute("top", "120px");
  spinner.setStyleAttribute("left", "220px");
  var refreshOpacityHandler = app.createClientHandler().forTargets(panel).setStyleAttribute('opacity', '0.5').forTargets(spinner).setVisible(true);
  var sheetId = formSheet.getSheetId();
  var hiddenSheetIdBox = app.createTextBox().setValue(sheetId).setId('sheetId').setName('sheetId').setVisible(false);  
  var hiddenColNumBox = app.createTextBox().setValue(columns).setId('numCols').setName('numCols').setVisible(false);
  panel.add(hiddenSheetIdBox);
  panel.add(hiddenColNumBox);
  var noDataLabel = app.createLabel("No data in sheet").setId('noDataLabel');
  noDataLabel.setVisible(false);
  panel.add(noDataLabel);
  panel.add(grid);
  outerScroll.add(panel);
  app.add(outerScroll);
  app.add(spinner);
  formulaCaddy_returnSheetUi(formSheet, properties);
  ss.show(app);        
}


function formulaCaddy_returnSheetUi(sheet, properties) {
  var allExcludedHeaders = getAllExludedHeaders();
  var alphabet = returnAlphabet();
  var app = UiApp.getActiveApplication();
  var panel = app.getElementById('panel');
  panel.setStyleAttribute('opacity','1');
  var spinner = app.getElementById('spinner');
  spinner.setVisible(false);
  var scrollPanel = app.createScrollPanel().setWidth("660px").setStyleAttribute('margin', '10px');
  var grid = app.getElementById('grid');
  var sheetId = sheet.getSheetId();
  var sheetProperties = properties;
  var hiddenSheetIdBox = app.getElementById('sheetId').setValue(sheetId);
  var columns = sheet.getLastColumn();
  var hiddenColNumBox = app.getElementById('numCols').setValue(columns);
  if (sheetProperties['formulaRow']) {
    var row = parseInt(sheetProperties['formulaRow']);
  } else {
    var row = 2;
  }
  var rowsArray = [2,3,4];
  var noDataLabel = app.getElementById('noDataLabel');
  if (columns>0) {
    var colValues = sheetProperties['colValues'];
    if (colValues) {
      colValues = Utilities.jsonParse(colValues);
    } else {
      colValues = new Object();
    }
    noDataLabel.setVisible(false);
    grid.resize(4, columns+1);
    grid.setVisible(true);
    var horizontalPanel = app.createHorizontalPanel().setId('horizPanel');
    var formulaRowLabel = app.createLabel("Row containing master values/formulas");
    var formulaRowSelect = app.createListBox().setId('formulaRowSelect').setName("formulaRow");
    var rowSelectHandler = app.createServerHandler('formulaCaddy_refreshRowFormulas').addCallbackElement(panel);
    var lastRow = sheet.getLastRow();
    for (var i=2; i<lastRow+1; i++) {
      formulaRowSelect.addItem(i);
    }
    var rowSelectIndex = rowsArray.indexOf(row);
    formulaRowSelect.setSelectedIndex(rowSelectIndex);
    grid.setBorderWidth(1).setCellSpacing(0).setStyleAttribute('borderColor','#E5E5E5').setStyleAttribute('opacity', '1');
    formulaRowSelect.addChangeHandler(rowSelectHandler);
    horizontalPanel.add(formulaRowLabel);
    horizontalPanel.add(formulaRowSelect);
    horizontalPanel.add(formulaRowLabel);
    panel.add(horizontalPanel);
   
    var headerRange = sheet.getRange(1,1,1,columns);
    var headers = headerRange.getValues()[0];
    var headerBgs = headerRange.getBackgroundColors()[0];
    grid.setWidget(0, 0, app.createLabel("Column"));
    grid.setWidget(1, 0, app.createLabel("Header"));  
    grid.setWidget(2, 0, app.createLabel("Value/Formula"));  
    grid.setWidget(3, 0, app.createLabel("Paste as values?"));  
    var onButtonHandlers = [];
    var onButtonServerHandlers = [];
    var offButtonHandlers = [];
    var offButtonServerHandlers = [];
    var onButtons = [];
    var offButtons = [];
    var buttonValues = [];
    var formulaLabels = [];
    var asValuesCheckBoxes = [];
    for (var i=0; i<columns; i++) {
      onButtons[i] = app.createButton(alphabet[i]).setId('onButton-'+sheetId+'-'+i).setStyleAttribute('background', 'whiteSmoke').setWidth("50px");
      offButtons[i] = app.createButton(alphabet[i]).setId('offButton-'+sheetId+'-'+i).setStyleAttribute('background', '#E5E5E5').setStyleAttribute('border', '2px solid grey').setWidth("50px").setVisible(false);
      buttonValues[i] = app.createTextBox().setVisible(false).setText(i+'-off').setName('bv-'+sheetId+'-'+i);
      var buttonPanel = app.createHorizontalPanel();
      buttonPanel.add(onButtons[i])
                 .add(offButtons[i])
                 .add(buttonValues[i])
                 .setStyleAttribute('width',"80px")
                 .setHorizontalAlignment(UiApp.HorizontalAlignment.CENTER);
      var formulas = formulaCaddy_getSheetFormulas(sheet, row);
      var formulaLabel = app.createLabel(formulas[i]).setId('formula-'+i).setStyleAttribute('opacity', '0.5'); 
      asValuesCheckBoxes[i] = app.createCheckBox().setId('asValues-'+sheetId+'-'+i).setName('av-'+sheetId+'-'+i).setEnabled(false).setValue(false);
      grid.setWidget(0, i+1, buttonPanel).setStyleAttribute(0, i+1, 'backgroundColor', 'whiteSmoke').setStyleAttribute(0, i+1, 'textAlign', 'center');
      grid.setWidget(1, i+1, app.createLabel(headers[i]));
      grid.setWidget(2, i+1, formulaLabel);   
      grid.setWidget(3, i+1, asValuesCheckBoxes[i]);
      if ((headerBgs[i] == "#DDDDDD")||(allExcludedHeaders.indexOf(headers[i])!=-1)) {
        grid.setStyleAttribute(1, i+1, 'backgroundColor', '#DDDDDD');
        onButtons[i].setEnabled(false).setVisible(false);
        offButtons[i].setEnabled(false).setVisible(false);
        asValuesCheckBoxes[i].setEnabled(false).setVisible(false);
      } else { 
        onButtonHandlers[i] = app.createClientHandler().forTargets(onButtons[i]).setVisible(false).forTargets(offButtons[i]).setVisible(true).forTargets(asValuesCheckBoxes[i]).setEnabled(true).forTargets(buttonValues[i]).setText(i+'-on');
        onButtonServerHandlers[i] = app.createServerHandler('toggleOpacity').addCallbackElement(panel);
        offButtonHandlers[i] = app.createClientHandler().forTargets(offButtons[i]).setVisible(false).forTargets(onButtons[i]).setVisible(true).forTargets(asValuesCheckBoxes[i]).setEnabled(false).setValue(false).forTargets(buttonValues[i]).setText(i+'-off');
        offButtonServerHandlers[i] = app.createServerHandler('toggleOpacity').addCallbackElement(panel);
        onButtons[i].addClickHandler(onButtonHandlers[i]).addClickHandler(onButtonServerHandlers[i]);
        offButtons[i].addClickHandler(offButtonHandlers[i]).addClickHandler(offButtonServerHandlers[i]);
      }
      if ((colValues['col-'+(i+1)])&&(allExcludedHeaders.indexOf(headers[i])==-1)) {
        onButtons[i].setVisible(false);
        offButtons[i].setVisible(true);
        buttonValues[i].setText(i+'-on');
        formulaLabel.setStyleAttribute('opacity','1');
        asValuesCheckBoxes[i].setEnabled(true);
        if (colValues['col-'+(i+1)]=='1') {
          asValuesCheckBoxes[i].setValue(true);
        }
      } else {
        asValuesCheckBoxes[i].setValue(false);
      }
    }
    scrollPanel.add(grid);
    panel.add(scrollPanel);
    var saveHandler = app.createServerHandler('manualSave').addCallbackElement(panel);
    var saveClientHandler = app.createClientHandler().forTargets(spinner).setVisible(true).forTargets(panel).setStyleAttribute('opacity', '0.5');
    var button = app.createButton("Save settings").setId('button').addClickHandler(saveHandler).addClickHandler(saveClientHandler);
    panel.add(button);
  } else {
    noDataLabel.setVisible(true);
    var saveHandler = app.createServerHandler('manualSave').addCallbackElement(panel);
    var button = app.createButton("Save settings").setId('button').addClickHandler(saveHandler).setVisible(false);
    panel.add(button);
  }
  return app;
}



function manualSave(e) {
  var app = UiApp.getActiveApplication();
  saveformulaCaddySettings(e);
  app.close();
  return app;
}



function waitingIcon() {
  var app = UiApp.createApplication().setHeight(250).setWidth(200);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var waitingImageUrl = this.MULEICONURL;
  var image = app.createImage(waitingImageUrl).setWidth("125px").setStyleAttribute('marginLeft', '25px');
  app.add(image);
  app.add(app.createLabel('Please be patient as formulaCaddy formulas are recalculated and pasted down their designated columns. For complex spreadsheets this can take some time.'));
  ss.show(app);
  return app;
}

function closeIcon(e) {
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}


function formulaCaddy_getSheetFormulas(sheet, row) {
  var columns = sheet.getLastColumn();
  var range = sheet.getRange(row, 1, 1, columns)
  var formulas = range.getFormulas()[0];
  var values = range.getValues()[0];
  for (var i=0; i<formulas.length; i++) {
    if (formulas[i]=='') {
      if (typeof values[i] == 'string') {
        formulas[i]=values[i].substring(0,15);
      }  else {
        formulas[i]=values[i];
      }
    } else {
      formulas[i] = formulas[i].substring(0,15);
    }
    if (formulas[i].length==15) {
      formulas[i] += "...";
    }
  }
  return formulas;
}
  

function formulaCaddy_refreshRowFormulas(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetProperties = ScriptProperties.getProperties();
  var sheetId = e.parameter.sheetId;
  var colValues = sheetProperties['colValues'];
  if (colValues) {
    colValues = Utilities.jsonParse(colValues);
  } else {
    colValues = new Object();
  }
  var app = UiApp.getActiveApplication();
  var grid = app.getElementById('grid');
  var row = e.parameter["formulaRow"];
  var sheet = getSheetById(ss, sheetId);
  var formulas = formulaCaddy_getSheetFormulas(sheet, row);
  
  for (var i=0; i<formulas.length; i++) {
    var formulaLabel = app.createLabel(formulas[i]).setId('formula-'+i).setStyleAttribute('opacity', '0.35');
    if (colValues['col-'+(i+1)]) {
      formulaLabel.setStyleAttribute('opacity', '1');
    }
    grid.setWidget(2, i+1, formulaLabel); 
  }
  return app;
}



function toggleOpacity(e) {
  var app = UiApp.getActiveApplication();
  var sheetId = e.parameter.sheetId;
  var numCols = e.parameter.numCols;
  for (var i=0; i<numCols; i++ ) {
  var buttonValue = e.parameter['bv-'+sheetId+'-'+i];
  if (buttonValue) {
    buttonValue = buttonValue.split("-");
  }
  var label = app.getElementById('formula-'+buttonValue[0]);
  if (buttonValue[1] == "on") {
    label.setStyleAttribute('opacity','1');
  } else {
    label.setStyleAttribute('opacity','0.5');
  }
}
  return app;
}


function saveformulaCaddySettings(e) {
  var app = UiApp.getActiveApplication();
  var properties = ScriptProperties.getProperties();
  var sheetId = e.parameter.sheetId;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getSheetById(ss, sheetId);
  var numCols = e.parameter.numCols;
  var formulaRow = e.parameter['formulaRow']
  properties["formulaRow"] = formulaRow;
  var copyDownFormulas = "false";
  var colValues = {};
    for (var j=0; j<numCols; j++) {
      var buttonValue = e.parameter['bv-'+sheetId+'-'+j];
      if (buttonValue==j+'-on') {
        copyDownFormulas = "true";
        var asValuesOption = e.parameter['av-'+sheetId+'-'+j];
        if (asValuesOption == 'false') {
          colValues["col-" + (j+1)] = 2;
        } else {
          colValues["col-" + (j+1)] = 1; 
        }
        properties["formula-" + (j+1)] = sheet.getRange(formulaRow, j+1).getFormula().toString();
      } else {
        try {
          delete colValues["col-" + (j+1)];
          delete properties["formula-" + (j+1)];
          ScriptProperties.deleteProperty("formula-" + (j+1));
        } catch(err) { 
        }
      }
    }
  properties.copyDownFormulas = copyDownFormulas;
  properties.colValues = Utilities.jsonStringify(colValues);
  ScriptProperties.setProperties(properties);
  returnCopydownStatusColIndex();
  app.close();
  return app;
}
function formMule_preconfig() {
// if you are interested in sharing your complete workflow system for others to copy (with script settings)
// Select the "Generate preconfig()" option in the menu and
//#######Paste preconfiguration code below before sharing your system for copy#######


  
  
  
    
//#######End preconfiguration code#######
// Note: By design, calendarID's will not be copied along with script settings. 
 //Fetch system name, if this script is part of a New Visions system
  var systemName = getSystemName();
  if (systemName) {
    ScriptProperties.setProperty('systemName', systemName)
  }
  //Fetch institutional tracking code.  If it exists, launch initialize function (autolaunch step 1 for repeat users)
  //If it doesn't exist, the checkInstitutionalTrackingCode() will launch the tracking settings UI.
  var institutionalTrackingString = checkInstitutionalTrackingCode();
}
function formMule_extractorWindow () {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  var propertyString = '';
  var excludeProperties = ['formmule_sid', 'caseNo', 'ssId', 'SSId', 'calendarToken', 'webAppUrl','twilioNumber','accountSID','lastPhone','authToken','preconfigStatus', 'ssId', 'ssKey', 'formulaTriggerSet'];
  for (var key in properties) {
    if (excludeProperties.indexOf(key)==-1) {
      var keyProperty = properties[key]; //.replace(/[/\\*]/g, "\\\\");                                     
      propertyString += "   ScriptProperties.setProperty('" + key + "','" + keyProperty + "');\n";
    }
  }
  var app = UiApp.createApplication().setHeight(500).setTitle("Export preconfig() settings");
  var panel = app.createVerticalPanel().setWidth("100%").setHeight("100%");
  var labelText = "Copying a Google Spreadsheet copies scripts along with it, but without any of the script settings saved.  This normally makes it hard to share full, script-enabled Spreadsheet systems. ";
  labelText += " You can solve this problem by pasting the code below into a script file called \"paste preconfig here\" (go to Script Editor and look in left sidebar) prior to publishing your Spreadsheet for others to copy. \n";
  labelText += " After a user copies your spreadsheet, they will select \"Run initial installation.\"  This will preconfigure all needed script settings.  If you got this workflow from someone as a copy of a spreadsheet, this has probably already been done for you.";
  var label = app.createLabel(labelText);
  var window = app.createTextArea();
  var codeString = "//This section sets all script properties associated with this formMule profile \n";
  codeString += "var preconfigStatus = ScriptProperties.getProperty('preconfigStatus');\n";
  codeString += "if (preconfigStatus!='true') {\n";
  codeString += propertyString; 
  codeString += "};\n";
  codeString += "ScriptProperties.setProperty('preconfigStatus','true');\n";
  codeString += "var ss = SpreadsheetApp.getActiveSpreadsheet();\n";
  if (properties.formulaTriggerSet == "true") {
    codeString += "setCopyDownTrigger(); \n";
  }
  
 //generate msgbox warning code if automated email or calendar is enabled in template 
  if ((properties.calendarStatus == 'true')||(properties.emailStatus == 'true')) {
    codeString += "\n \n";
    codeString += "//Custom popup and function calls to prompt user for additional settings \n";
    if (ScriptProperties.getProperty('emailStatus')=="true") {
       codeString += "Browser.msgBox(\"You will want to check the email template sheets to ensure the correct sender and recipients are listed before using.\");\n";
    }
    if (ScriptProperties.getProperty('calendarStatus')=="true") {
    codeString += "Browser.msgBox(\"You need to set a new calendarID for this script before it will work.\");\n";
    codeString += "formMule_setCalendarSettings();";
    }
  }
  codeString += "ss.toast('Custom formMule preconfiguration ran successfully. Please check formMule menu options to confirm system settings.');\n";
  window.setText(codeString).setWidth("100%").setHeight("400px");
  app.add(label);
  panel.add(window);
  app.add(panel);
  ss.show(app);
  return app;
}
function formMule_loadingPreview() {
  var app = UiApp.createApplication().setWidth(140).setHeight(140);
  var label = app.createLabel("Loading preview...").setStyleAttribute('textAlign', 'center');
  var image = app.createImage(MULEICONURL).setHeight("100px").setStyleAttribute('marginLeft', '20px');
  app.add(label);
  app.add(image);
  ss.show(app);
  formMule_previewSend();
  return app;
}


function formMule_previewSend() {
  var app = UiApp.getActiveApplication();
  var properties = ScriptProperties.getProperties();
  app.close();
  var manual = true;
  var app = UiApp.createApplication().setTitle("Here's what your merge will look like...").setWidth(590).setHeight(420);
  var panel = app.createVerticalPanel();
  var tabPanel = app.createTabPanel().setHeight("330px");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = properties.sheetName;
  var sheet = ss.getSheetByName(sheetName);
  var copyDownOption = properties.copyDownOption;
  var lastRow = sheet.getLastRow();
  var emailData = new Array();
  var calData = new Array();
  var smsData = new Array();
  var vmData = new Array();
  var calUpData = new Array();
  var oldCalData = new Array();
  var k = 0;
  var m = 0;
  var n = 0;
  var s = 0;
  var v = 0;
  var calendarStatus = properties.calendarStatus;
  var calendarUpdateStatus = properties.calendarUpdateStatus;
  var calTrigger = properties.calTrigger;
  var calUpdateTrigger = properties.calUpdateTrigger;
  var userEmail = Session.getEffectiveUser().getEmail();
  if (((calendarStatus == "true")&&(manual==true))||((calendarStatus == "true")&&(calTrigger=="true"))||((calendarUpdateStatus == "true")&&(manual==true))||((calendarUpdateStatus == "true")&&(calUpdateTrigger=="true"))) {
    var eventIdCol = properties.eventIdCol;
    var calConditionsString = properties.calendarConditions;
    var calCondObject = Utilities.jsonParse(calConditionsString);
    var calUpdateConditionsString = properties.calendarUpdateConditions;
    var calUpdateCondObject = Utilities.jsonParse(calUpdateConditionsString);
    var calDeleteConditionsString = properties.calendarDeleteConditions;
    var calDeleteCondObject = Utilities.jsonParse(calDeleteConditionsString);
    var calendarToken  = properties.calendarToken;
    var eventTitleToken = properties.eventTitleToken;
    var locationToken = properties.locationToken;
    var guestsToken = properties.guests;
    var emailInvites = properties.emailInvites;
    if (emailInvites == "true") { emailInvites = true; } else { emailInvites = false;}
    var allDay = properties.allDay;
    var startTimeToken = properties.startTimeToken;
    var endTimeToken = properties.endTimeToken;
    var descToken = properties.descToken;
    var reminderType = properties.reminderType;
    var minBefore = properties.minBefore;
    var calRepeats = properties.calendarWeeklyRepeats;
    var calWeekdays = properties.calendarWeekdays;
  }
  var emailStatus = properties.emailStatus;
  var emailTrigger = properties.emailTrigger;
  if (((emailStatus=="true")&&(manual==true))||((emailStatus=="true")&&(emailTrigger=="true"))) {
    var emailCondString = properties.emailConditions;
    var emailCondObject = Utilities.jsonParse(emailCondString);
    var numSelected = properties.numSelected;
    var templateSheetNames = [];
    var normalizedSendColNames = [];
    var sendersTemplates = [];
    var recipientsTemplates = [];
    var ccRecipientsTemplates = [];
    var subjectTemplates = [];
    var bodyTemplates = [];  
    var langTemplates = [];
    for (var i=0; i<numSelected; i++) {
       var templateSheet = ss.getSheetByName(emailCondObject['sht-'+i]);
       if (!templateSheet) {
         Browser.msgBox("An error occurred: It appears one of your email template sheets was deleted or renamed. Go back to step 2a to fix this.");
         app.close();
         return app;
       }
       normalizedSendColNames[i] = formMule_normalizeHeader(emailCondObject['sht-'+i] + " Status");
       templateSheetNames[i] = templateSheet.getName();
       sendersTemplates[i] = templateSheet.getRange("B1").getValue();
       recipientsTemplates[i] = templateSheet.getRange("B2").getValue();
       ccRecipientsTemplates[i] = templateSheet.getRange("B3").getValue();
       subjectTemplates[i] = templateSheet.getRange("B4").getValue();
       bodyTemplates[i] = templateSheet.getRange("B5").getValue();
       langTemplates[i] = templateSheet.getRange("B6").getValue();
    }
  } 
  var smsStatus = properties.smsEnabled;
  var smsTrigger = properties.smsTrigger;
  var vmStatus = properties.vmEnabled;
  var vmTrigger = properties.vmTrigger;
  var accountSID = properties.accountSID;
  var authToken = properties.authToken;
  var maxTexts = properties.smsMaxLength;
  var dataSheetName = properties.sheetName;
  var dataSheet = ss.getSheetByName(dataSheetName);
  var headers = formMule_fetchHeaders(dataSheet);
  var eventIdIndex = headers.indexOf("Event Id");
  var updateEventIdIndex = headers.indexOf(eventIdCol);
  if (eventIdCol) {
    eventIdCol = formMule_normalizeHeader(eventIdCol);
  }
  var z=2;
  if (copyDownOption=="true") {
     z=3;
  }
  var dataRange = dataSheet.getRange(z, 1, dataSheet.getLastRow()-(z-1), dataSheet.getLastColumn());
  var lastHeader = dataSheet.getRange(1, dataSheet.getLastColumn());
  var lastHeaderValue = formMule_normalizeHeader(lastHeader.getValue());  
  var test = dataRange.getValues()
  // Create one JavaScript object per row of data.
  var objects = formMule_getRowsData(dataSheet, dataRange, 1);
  // For every row object, create a personalized email from a template and send
  // it to the appropriate person.
  for (var j = 0; j < objects.length; ++j) {
    // Get a row object
    var rowData = objects[j];
    //Get merge and update status for the row
    var confirmation = '';
    var calUpdateConfirmation = '';
    var found = '';
    
    var calConditionTest = false;
    var calUpdateConditionTest = false;
    var calDeleteConditionTest = false;
    if (calendarStatus=="true") {
      //test calendar event create condition
      if ((calConditionsString)&&(calConditionsString!='')) {
        var calConditionTest = formMule_evaluateConditions(calCondObject, 0, rowData);
      } 
    }
    
    if (calendarUpdateStatus=="true") {
      var calUpdateStatus = rowData.eventUpdateStatus;
      if (calUpdateStatus=='') {
        //test calendar event update condition
        if ((calUpdateConditionsString)&&(calUpdateConditionsString!='')) {
          var calUpdateConditionTest = formMule_evaluateUpdateConditions(calUpdateCondObject, 0, rowData);    
        } 
        //test calendar event deletion condition
        if ((calDeleteConditionsString)&&(calDeleteConditionsString!='')) {
          var calDeleteConditionTest = formMule_evaluateUpdateConditions(calDeleteCondObject, 0, rowData);    
        } 
      }
    }
    
    if ((calendarToken)&&(calendarToken!='')&&((calUpdateConditionTest)||(calConditionTest)||(calDeleteConditionTest))&&(m<7)) {
      var calendarId  = formMule_fillInTemplateFromObject(calendarToken, rowData);
      var lastCalendarId = CacheService.getPrivateCache().get('lastCalendarId');
      if (lastCalendarId!=calendarId) {
        var calendar = CalendarApp.getCalendarById(calendarId);
      }
    }
    if ((!calendarToken)&&((calUpdateConditionTest)||(calConditionTest)||(calDeleteConditionTest))) {
      Browser.msgBox("You have not specified a calendar ID for your calendar merge.  Return to step 2b.");
          return;
    }
    // check for calendar status setting
    if (calConditionTest == true) {
      calData[m] = new Object();
      if (calendar) {
        calData[m].calendar = calendar.getName();
        var eventTitle = formMule_fillInTemplateFromObject(eventTitleToken, rowData);
        calData[m].eventTitle = eventTitle;
        var location = formMule_fillInTemplateFromObject(locationToken, rowData);
        calData[m].location = location;
        var desc = formMule_fillInTemplateFromObject(descToken, rowData);
        calData[m].desc = desc;
        var guests = formMule_fillInTemplateFromObject(guestsToken, rowData);
        calData[m].guests = guests;
        var startTimeStamp;
        var endTimeStamp;
        var timeZone = Session.getTimeZone();
      if (!(allDay=="true")) {
        var startTimeStamp = new Date(formMule_fillInTemplateFromObject(startTimeToken, rowData));
        endTimeStamp = new Date(formMule_fillInTemplateFromObject(endTimeToken, rowData));
        calData[m].allDay = allDay;
        calData[m].startTimeStamp = startTimeStamp;
        calData[m].endTimeStamp = endTimeStamp;
      } else {
        calData[m].allDay = allDay;
        startTimeStamp = new Date(formMule_fillInTemplateFromObject(startTimeToken, rowData));
        startTimeStamp = Utilities.formatDate(startTimeStamp, timeZone, "MM/dd/yyyy");
        calData[m].startTimeStamp = startTimeStamp;
        calData[m].endTimeStamp = '';
      }
      calData[m].emailInvites = emailInvites;
      calData[m].repeats = formMule_fillInTemplateFromObject(calRepeats, rowData);
      calData[m].weekdays = formMule_fillInTemplateFromObject(calWeekdays, rowData);
      } else {
        calData[m].calendar = "Calendar could not be found. Check your calendar Id";
      }
    m++;
  } //end calendar create condition test
  if ((calUpdateConditionTest == true)||(calDeleteConditionTest == true)) {  
    var eventId = rowData[eventIdCol];
    calUpData[n] = new Object();
    oldCalData[n] = new Object();
    if (n<7) {
    if (calendar) {
    calUpData[n].calendar = calendar.getName();
    var found = false;
    try {
      var future = new Date();
      future.setDate(future.getDate()+360);
      var past = new Date();
      past.setDate(past.getDate()-360);
      var events = calendar.getEvents(past, future);
      for (var g = 0; g<events.length; g++) {
        if (events[g].getId()==eventId) {
          var oldEvent = events[g];
          found = true;
        }
      }
    } catch(err) {
      found = false;
    }
    } else {
      Browser.msgBox("Calendar could not be found");
      return;
    }
    if ((found == true)&&((calUpdateConditionTest==true)||(calDeleteConditionTest==true))) {
      oldCalData[n].eventTitle = oldEvent.getTitle();
      oldCalData[n].location = oldEvent.getLocation(); 
      oldCalData[n].desc = oldEvent.getDescription();
      oldCalData[n].guests = oldEvent.getGuestList().join();
      if ((found==true)&&(calUpdateConditionTest == true)) {
        var eventTitle = formMule_fillInTemplateFromObject(eventTitleToken, rowData);
        calUpData[n].eventTitle = eventTitle;
        var location = formMule_fillInTemplateFromObject(locationToken, rowData);
        calUpData[n].location = location;
        var desc = formMule_fillInTemplateFromObject(descToken, rowData);
        calUpData[n].desc = desc;
        var guests = formMule_fillInTemplateFromObject(guestsToken, rowData);
        calUpData[n].guests = guests;
        var startTimeStamp;
        var endTimeStamp;
        var timeZone = Session.getTimeZone();
      if (!(allDay=="true")) {
        var startTimeStamp = new Date(formMule_fillInTemplateFromObject(startTimeToken, rowData));
        endTimeStamp = new Date(formMule_fillInTemplateFromObject(endTimeToken, rowData));
        calUpData[n].allDay = allDay;
        calUpData[n].startTimeStamp = startTimeStamp;
        calUpData[n].endTimeStamp = endTimeStamp;
      } else {
        calUpData[n].allDay = allDay;
        startTimeStamp = new Date(formMule_fillInTemplateFromObject(startTimeToken, rowData));
        startTimeStamp = Utilities.formatDate(startTimeStamp, timeZone, "MM/dd/yyyy");
        calUpData[n].startTimeStamp = startTimeStamp;
        calUpData[n].endTimeStamp = '';
      }
        calUpData[n].emailInvites = emailInvites;
      }      
    // same as above, but for old event settings
    var oldAllDay = oldEvent.isAllDayEvent();
    if (!(oldAllDay==true)) {
      var startTimeStamp = oldEvent.getStartTime();
      endTimeStamp = oldEvent.getEndTime();
      oldCalData[n].allDay = oldAllDay;
      oldCalData[n].startTimeStamp = startTimeStamp;
      oldCalData[n].endTimeStamp = endTimeStamp;
    } else {
      oldCalData[n].allDay = oldAllDay;
      startTimeStamp = oldEvent.getStartTime();
      startTimeStamp = Utilities.formatDate(startTimeStamp, timeZone, "MM/dd/yyyy");
      oldCalData[n].startTimeStamp = startTimeStamp;
      oldCalData[n].endTimeStamp = '';
    }
      oldCalData[n].emailInvites = oldEvent.getEmailReminders(); 
    }
     if ((found==true)&&(calDeleteConditionTest == true)) {
      calUpData[n].eventTitle = "Event will be deleted";
      calUpData[n].location = "Event will be deleted";
      calUpData[n].desc = "Event will be deleted";
      calUpData[n].guests = "Event will be deleted";
      calUpData[n].allDay = "Event will be deleted";
      calUpData[n].startTimeStamp = "Event will be deleted";
      calUpData[n].endTimeStamp = "Event will be deleted";
      calUpData[n].emailInvites = "Event will be deleted";
    }
      
    if (found==false) {
      oldCalData[n].eventTitle = "Event " + eventId + " could not be found. Check your event id settings.";
    }
   }
    n++;
  } //end calendar update condition test
  // check for email status setting 
  if ((emailStatus=="true")&&(manual==true)) {
  for (var i=0; i<numSelected; i++) {
    if ((emailCondString)&&(emailCondString!='')) {
    var emailConditionTest = formMule_evaluateConditions(emailCondObject, i, rowData);
    }
    if ((emailConditionTest == true)||(!emailCondString)) {
    emailData[k] = new Object();
    emailData[k].num = 'Template used: '+templateSheetNames[i];
    if (k<7) {
      var sendersTemplate = sendersTemplates[i];
      var recipientsTemplate = recipientsTemplates[i];
      var ccRecipientsTemplate = ccRecipientsTemplates[i];
      var subjectTemplate = subjectTemplates[i];
      var bodyTemplate = bodyTemplates[i];
      var langTemplate = langTemplates[i];
      emailData[k].from = userEmail;
    // Generate a personalized email.
    // Given a template string, replace markers (for instance ${"First Name"}) with
    // the corresponding value in a row object (for instance rowData.firstName).
      var from = formMule_fillInTemplateFromObject(sendersTemplate, rowData);
      emailData[k].replyTo = from;
      var to = formMule_fillInTemplateFromObject(recipientsTemplate, rowData);
      try {
      if (to!='') {
        emailData[k].to = to;
        var cc = formMule_fillInTemplateFromObject(ccRecipientsTemplate, rowData);
        emailData[k].cc = cc;
        var subject = formMule_fillInTemplateFromObject(subjectTemplate, rowData);
        emailData[k].subject = subject;
        var body = formMule_fillInTemplateFromObject(bodyTemplate, rowData);
        var lang = formMule_fillInTemplateFromObject(langTemplate, rowData);
        if ((lang)&&(lang!='')) {
          var translation = LanguageApp.translate("Automated Translation", '', lang);
          var divider = '<h3>####### ' + translation + ' #######</h3>';
          try {
            body += divider + LanguageApp.translate(body, '', lang);
          } catch(err) {
            body = "A translation error occurred.  Check your language code(s)";
          }
        }
        emailData[k].body = body;
        var now = new Date();
        now = Utilities.formatDate(now, timeZone, "MM/dd/yy' at 'h:mm:ss a")
        if (cc!='') { var ccMsg = ", and cc to "+cc; } else {ccMsg = ''}
        if ((i>0)&&(i<numSelected-1)) { var addSemiColon = "; "} else { var addSemiColon = ""; }
        confirmation += " Will attempt to send Email"+(i+1)+" to "+to+ccMsg+addSemiColon;
        } else {
          emailData[k].to = "missing";
          emailData[k].cc = cc;
          emailData[k].subject = subject;
          emailData[k].body = body;
          confirmation += " Email"+(i+1)+" error: Template missing \"To\" address";
          var error = true;
        }
      } catch(err) {
        confirmation += " Email"+(i+1)+" error: " + err;
        var error = true;
      }
    }
      k++;
     } // end per email conditional check
     } // end i loop through email templates
    } // end conditional test for email
      //begin SMS section
    if (((smsStatus=="true")&&(manual==true))||((smsStatus=="true")&&(smsTrigger=="true"))) { 
      var twilioNumber = properties.twilioNumber;
      var smsPropertyString = properties.smsPropertyString;
      for (var i=0; i<properties.smsNumSelected; i++) {
        if ((smsPropertyString)&&(emailCondString!='')) {
          var smsPropertyObject = Utilities.jsonParse(smsPropertyString);
          var smsConditionTest = formMule_evaluateSMSConditions(smsPropertyObject, i, rowData);  
          if ((smsConditionTest == true)||(!smsPropertyString)) {
            smsData[s] = new Object();
            smsData[s].phoneNumber = formMule_fillInTemplateFromObject(smsPropertyObject['smsPhone-'+i], rowData);
            var lang = formMule_fillInTemplateFromObject(smsPropertyObject['smsLang-'+i], rowData);
            lang = lang.trim();
            var body = formMule_fillInTemplateFromObject(smsPropertyObject['smsBody-'+i], rowData, true);
            try {
              if ((lang) && (lang!='')) {
                body = LanguageApp.translate(body, '', lang);
              }
            } catch(err) {
              body = 'A translation error occurred. Check your language code(s).';
            }
            smsData[s].body = formMule_splitTexts(body, maxTexts).join(" || ");
            smsData[s].num = 'Template used: '+smsPropertyObject['smsName-'+i];
            s++;
          }
        }
      }
    }
    //end SMS Section
    //begin VM section
    if (((vmStatus=="true")&&(manual==true))||((vmStatus=="true")&&(vmTrigger=="true"))) {
      var vmPropertyString = properties.vmPropertyString;
      for (var i=0; i<properties.vmNumSelected; i++) {
        if ((vmPropertyString)&&(emailCondString!='')) {
          var vmPropertyObject = Utilities.jsonParse(vmPropertyString);
          var vmConditionTest = formMule_evaluateVMConditions(vmPropertyObject, i, rowData);  
          if ((vmConditionTest == true)||(!vmPropertyString)) {
            vmData[v] = new Object();
            vmData[v].phoneNumber = formMule_fillInTemplateFromObject(vmPropertyObject['vmPhone-'+i], rowData);
            var lang = formMule_fillInTemplateFromObject(vmPropertyObject['vmLang-'+i], rowData);
            lang = lang.trim();
            var body = formMule_fillInTemplateFromObject(vmPropertyObject['vmBody-'+i], rowData, true);
            try {
              if ((lang) && (lang!='')) {
                body = LanguageApp.translate(body, '', lang);
              }
            } catch(err) {
              body = 'A translation error occurred. Check your language code(s).';
            }
            vmData[v].body = body;
            if (vmPropertyObject['vmRecordOption-'+i]=="Yes") {
              vmData[v].RequestResponse = "Yes";
            } else {
              vmData[v].RequestResponse = "No";
            }
            if ((vmPropertyObject['vmSoundFile-'+i])&&(vmPropertyObject['vmSoundFile-'+i])!='') {  
              vmData[v].SoundFile = vmPropertyObject['vmSoundFile-'+i];
            }
            v++;
          }
        }
      }
    }     
  } //end j loop through spreadsheet rows
  
  //begin email status panel
  var emailStatus = ScriptProperties.getProperty("emailStatus");
  var emailPanel = app.createVerticalPanel().setHeight("270px");
  emailPanel.add(app.createLabel("Email merge").setStyleAttribute("fontSize","16px").setStyleAttribute("borderTop", "1px solid #EBEBEB").setStyleAttribute("borderBottom", "1px solid #EBEBEB").setStyleAttribute("padding","5px").setWidth("100%").setStyleAttribute("backgroundColor", "whiteSmoke"));
  if (emailStatus=="true") {
  var emailScrollPanel = app.createScrollPanel().setWidth("560px").setHeight("260px");
  var emailVerticalPanel = app.createVerticalPanel().setWidth("100%");
  if (emailData.length<6) {
    var previewNum = emailData.length;
  } else {
    var previewNum = 6
  }
  if (previewNum>0) {
  for (var i=0; i<previewNum; i++) {
    var emailLabel = app.createLabel("Email #"+(i+1)+" - "+emailData[i].num+" Template").setWidth("560px").setStyleAttribute('backgroundColor', '#E5E5E5').setStyleAttribute('textAlign', 'center');
    var grid2 = app.createGrid(6, 2).setBorderWidth(0).setCellSpacing(4).setWidth("100%");
    grid2.setWidget(1, 0, app.createLabel('From:')).setStyleAttribute(1, 0, "width", "100px").setStyleAttribute(1,0,'backgroundColor', 'whiteSmoke');
    grid2.setWidget(2, 0, app.createLabel('Reply to:')).setStyleAttribute(2,0,'backgroundColor', 'whiteSmoke');
    grid2.setWidget(0, 0, app.createLabel('To:')).setStyleAttribute(0,0,'backgroundColor', 'whiteSmoke');
    grid2.setWidget(3, 0, app.createLabel('CC:')).setStyleAttribute(3,0,'backgroundColor', 'whiteSmoke');
    grid2.setWidget(4, 0, app.createLabel('Subject:')).setStyleAttribute(4,0,'backgroundColor', 'whiteSmoke');
    grid2.setWidget(5, 0, app.createLabel('Body:')).setStyleAttribute(5,0,'backgroundColor', 'whiteSmoke');
    grid2.setWidget(1, 1, app.createLabel(emailData[i].from)).setStyleAttribute(1,1,'backgroundColor', 'whiteSmoke');
    grid2.setWidget(2, 1, app.createLabel(emailData[i].replyTo)).setStyleAttribute(2,1,'backgroundColor', 'whiteSmoke');
    grid2.setWidget(0, 1, app.createLabel(emailData[i].to)).setStyleAttribute(0,1,'backgroundColor', 'whiteSmoke');
    grid2.setWidget(3, 1, app.createLabel(emailData[i].cc)).setStyleAttribute(3,1,'backgroundColor', 'whiteSmoke');
    grid2.setWidget(4, 1, app.createLabel(emailData[i].subject)).setStyleAttribute(4,1,'backgroundColor', 'whiteSmoke');
    grid2.setWidget(5, 1, app.createHTML(emailData[i].body)).setStyleAttribute(5,1,'backgroundColor', 'whiteSmoke');
    emailVerticalPanel.add(emailLabel);
    emailVerticalPanel.add(grid2);
    }
    var label2 = app.createLabel('Check out the first ' + previewNum + ' of the '+ emailData.length +' emails that will be sent.').setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px');
    emailPanel.add(label2);
   } else {
    var label2 = app.createLabel('No emails will be sent given your current settings.  This could be because of your email conditions (see Step 2a) or because all your data rows already have a status message in the \'Status\' column for your templates.').setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px');
    emailPanel.add(label2);
   }
    emailScrollPanel.add(emailVerticalPanel);
    emailPanel.add(emailScrollPanel);
 
  
  if (previewNum>0) {
    var label3 = app.createLabel("Note: Sender address is always that of the script installer. Images and links in the email body currently don\'t preview correctly.  We're working on it;)").setStyleAttribute('color', 'blue');
    emailPanel.add(label3);
  }
  } else {
    emailPanel.add(app.createLabel("Email merge is currently disabled. Want merged emails? You can enable this service in Step 2a.").setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px'));
  }
  tabPanel.add(emailPanel, "Emails");
  
  //begin calendar tab
  var calendarPanel = app.createVerticalPanel().setHeight("270px");
  var calendarStatus = ScriptProperties.getProperty("calendarStatus");
    calendarPanel.add(app.createLabel("Calendar Event Merge").setStyleAttribute("fontSize","16px").setStyleAttribute("borderTop", "1px solid #EBEBEB").setStyleAttribute("borderBottom", "1px solid #EBEBEB").setStyleAttribute("padding","5px").setWidth("100%").setStyleAttribute("backgroundColor", "whiteSmoke"));
  if (calendarStatus=="true") {
  var calScrollPanel = app.createScrollPanel().setWidth("560px").setHeight("260px");
  var calVerticalPanel = app.createVerticalPanel().setWidth("560px");
  if (calData.length<6) {
    var previewCalNum = calData.length;
  } else {
    var previewCalNum = 6
  }
  if (previewCalNum>0) {
  var label4 = app.createLabel('Check out the first ' + previewCalNum + ' of the '+ calData.length +' calendar events that will be created').setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px');
  calendarPanel.add(label4);
  calScrollPanel.add(calVerticalPanel);
  calendarPanel.add(calScrollPanel);
  for (var i=0; i<previewCalNum; i++) {
    var calLabel = app.createLabel("Calendar Event #"+(i+1)).setWidth("560px").setStyleAttribute('backgroundColor', '#E5E5E5').setStyleAttribute('textAlign', 'center');
    var grid3 = app.createGrid(10, 2).setBorderWidth(0).setCellSpacing(4).setWidth("100%");
    grid3.setWidget(0, 0, app.createLabel('Calendar:')).setStyleAttribute(0, 0, "width", "100px").setStyleAttribute(0,0,'backgroundColor', 'whiteSmoke');
    grid3.setWidget(1, 0, app.createLabel('Event title:')).setStyleAttribute(1,0,'backgroundColor', 'whiteSmoke');
    grid3.setWidget(2, 0, app.createLabel('Event type:')).setStyleAttribute(2,0,'backgroundColor', 'whiteSmoke');
    grid3.setWidget(3, 0, app.createLabel('Start time:')).setStyleAttribute(3,0,'backgroundColor', 'whiteSmoke');
    grid3.setWidget(4, 0, app.createLabel('End Time:')).setStyleAttribute(4,0,'backgroundColor', 'whiteSmoke');
    grid3.setWidget(5, 0, app.createLabel('Event Description:')).setStyleAttribute(5,0,'backgroundColor', 'whiteSmoke');
    grid3.setWidget(6, 0, app.createLabel('Guests:')).setStyleAttribute(6,0,'backgroundColor', 'whiteSmoke');
    grid3.setWidget(7, 0, app.createLabel('Email invites?')).setStyleAttribute(7,0,'backgroundColor', 'whiteSmoke');
    if (calData[i].repeats!='') {
      grid3.setWidget(8, 0, app.createLabel('Number of weeks to repeat')).setStyleAttribute(8,0,'backgroundColor', 'whiteSmoke');
    }
    if (calData[i].weekdays) {
      grid3.setWidget(9, 0, app.createLabel('Weekdays to repeat on')).setStyleAttribute(9,0,'backgroundColor', 'whiteSmoke');
    }
    grid3.setWidget(0, 1, app.createLabel(calData[i].calendar)).setStyleAttribute(0,1,'backgroundColor', 'whiteSmoke');
    grid3.setWidget(1, 1, app.createLabel(calData[i].eventTitle)).setStyleAttribute(1,1,'backgroundColor', 'whiteSmoke');
    if(calData[i].allDay=="true"){var text = "All Day";} else {var text = "Time slot"}
    grid3.setWidget(2, 1, app.createLabel(text)).setStyleAttribute(2,1,'backgroundColor', 'whiteSmoke');
    grid3.setWidget(3, 1, app.createLabel(calData[i].startTimeStamp)).setStyleAttribute(3,1,'backgroundColor', 'whiteSmoke');
    if (calData[i].allDay=="true"){var text = "n/a";} else {var text = calData[i].endTimeStamp;}
    grid3.setWidget(4, 1, app.createLabel(text)).setStyleAttribute(4,1,'backgroundColor', 'whiteSmoke');
    grid3.setWidget(5, 1, app.createLabel(calData[i].desc)).setStyleAttribute(5,1,'backgroundColor', 'whiteSmoke');
    grid3.setWidget(6, 1, app.createHTML(calData[i].guests)).setStyleAttribute(6,1,'backgroundColor', 'whiteSmoke');
    if(calData[i].emailInvites==true){var text = "Yes";} else {var text = "No"};
    grid3.setWidget(7, 1, app.createHTML(text)).setStyleAttribute(7,1,'backgroundColor', 'whiteSmoke');
    if (calData[i].repeats!='') {
      grid3.setWidget(8, 1, app.createHTML(calData[i].repeats)).setStyleAttribute(8,1,'backgroundColor', 'whiteSmoke');
    }
    if (calData[i].weekdays!='') {
      grid3.setWidget(9, 1, app.createHTML(calData[i].weekdays)).setStyleAttribute(9,1,'backgroundColor', 'whiteSmoke');
    }
    calVerticalPanel.add(calLabel);
    calVerticalPanel.add(grid3);
    }
  } else {
   var label5 = app.createLabel('No calendar appointments will be sent given your current settings.  This could be because of your calendar event creation condition setting (see Step 2b) or because all your data rows already have a status message in the \'Event Creation Status\' column.').setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px');
    calendarPanel.add(label5);
  }
  } else {
  calendarPanel.add(app.createLabel("Calendar merge is currently disabled. Want merged calendar events? You can enable this service in Step 2b.").setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px'));
  }
  tabPanel.add(calendarPanel, "Event Create")
  
  
  // Begin calendar update panel
  var calendarUpdatePanel = app.createVerticalPanel().setHeight("300px");
  var calendarStatus = ScriptProperties.getProperty("calendarUpdateStatus");
  calendarUpdatePanel.add(app.createLabel("Event Update Merge").setStyleAttribute("fontSize","16px").setStyleAttribute("borderTop", "1px solid #EBEBEB").setStyleAttribute("borderBottom", "1px solid #EBEBEB").setStyleAttribute("padding","5px").setWidth("100%").setStyleAttribute("backgroundColor", "whiteSmoke"));
  var calUpdateVerticalPanel = app.createVerticalPanel().setWidth("560px");
  if (calendarUpdateStatus=="true") {
  var calUpdateScrollPanel = app.createScrollPanel().setWidth("560px").setHeight("290px");
  if (calUpData.length<6) {
    var previewUpdateCalNum = calUpData.length;
  } else {
    var previewUpdateCalNum = 6
  }
  if (previewUpdateCalNum>0) {
  var label5 = app.createLabel('Check out the first ' + previewUpdateCalNum + ' of the '+ calUpData.length +' calendar events that will be updated').setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px');
  calendarUpdatePanel.add(label5);
  calUpdateScrollPanel.add(calUpdateVerticalPanel);
  calendarUpdatePanel.add(calUpdateScrollPanel);
  for (var i=0; i<previewUpdateCalNum; i++) {
    var calUpdateLabel = app.createLabel("Calendar Event Update #"+(i+1)).setWidth("560px").setStyleAttribute('backgroundColor', '#E5E5E5').setStyleAttribute('textAlign', 'center');
    var grid4 = app.createGrid(9, 3).setBorderWidth(0).setCellSpacing(4).setWidth("100%");
    grid4.setWidget(1, 0, app.createLabel('Calendar:')).setStyleAttribute(1, 0, "width", "100px").setStyleAttribute(1,0,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(2, 0, app.createLabel('Event title:')).setStyleAttribute(2, 0, "width", "100px").setStyleAttribute(2,0,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(3, 0, app.createLabel('Event type:')).setStyleAttribute(3,0,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(4, 0, app.createLabel('Start time:')).setStyleAttribute(4,0,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(5, 0, app.createLabel('End Time:')).setStyleAttribute(5,0,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(6, 0, app.createLabel('Event Description:')).setStyleAttribute(6,0,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(7, 0, app.createLabel('Guests:')).setStyleAttribute(7,0,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(8, 0, app.createLabel('Email invites?')).setStyleAttribute(8,0,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(0, 1, app.createLabel('Current Event Info'));
    grid4.setWidget(0, 2, app.createLabel('Updated Event Info'));
    grid4.setWidget(1, 1, app.createLabel(calUpData[i].calendar)).setStyleAttribute(1,1,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(2, 1, app.createLabel(oldCalData[i].eventTitle)).setStyleAttribute(2,1,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(2, 2, app.createLabel(calUpData[i].eventTitle)).setStyleAttribute(2,2,'backgroundColor', 'whiteSmoke');
    if(oldCalData[i].allDay=="true"){var oldText = "All Day";} else {var oldText = "Time slot"}
    grid4.setWidget(3, 1, app.createLabel(oldText)).setStyleAttribute(3,1,'backgroundColor', 'whiteSmoke');
    if(calUpData[i].allDay=="true"){var newText = "All Day";} else {var newText = "Time slot"}
    grid4.setWidget(3, 2, app.createLabel(newText)).setStyleAttribute(3,2,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(4, 1, app.createLabel(oldCalData[i].startTimeStamp)).setStyleAttribute(4,1,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(4, 2, app.createLabel(calUpData[i].startTimeStamp)).setStyleAttribute(4,2,'backgroundColor', 'whiteSmoke');
    if (oldCalData[i].allDay=="true"){var text = "n/a";} else {var text = oldCalData[i].endTimeStamp;} 
    if (calUpData[i].allDay=="true"){var newText = "n/a";} else {var newText = calUpData[i].endTimeStamp;}
    grid4.setWidget(5, 1, app.createLabel(text)).setStyleAttribute(5,1,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(5, 2, app.createLabel(newText)).setStyleAttribute(5,2,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(6, 1, app.createLabel(oldCalData[i].desc)).setStyleAttribute(6,1,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(6, 2, app.createLabel(calUpData[i].desc)).setStyleAttribute(6,2,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(7, 1, app.createHTML(oldCalData[i].guests)).setStyleAttribute(7,1,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(7, 2, app.createHTML(calUpData[i].guests)).setStyleAttribute(7,2,'backgroundColor', 'whiteSmoke');
    if(oldCalData[i].emailInvites=="true"){var text = "Yes";} else {var text = "No"};
    grid4.setWidget(8, 1, app.createHTML(text)).setStyleAttribute(8,1,'backgroundColor', 'whiteSmoke');
    grid4.setWidget(8, 2, app.createHTML("Currently not possible")).setStyleAttribute(8,2,'backgroundColor', 'whiteSmoke');
    calUpdateVerticalPanel.add(calUpdateLabel);
    calUpdateVerticalPanel.add(grid4);
    }
  } else {
   var label6 = app.createLabel('No calendar appointments will be updated given your current settings. This could be because of your update condition setting (see Step 2b) or because all your data rows already have a status message in the \'Event Update Status\' column.').setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px');
    calendarUpdatePanel.add(label6);
  }
  } else {
  calendarUpdatePanel.add(app.createLabel("The calendar update feature is currently disabled. Want to update merged calendar events? You can enable this service in Step 2b.").setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px'));
  }
  tabPanel.add(calendarUpdatePanel, "Event Update");
  //  End calendar update panel
  
  //begin SMS panel
  var smsStatus = properties.smsEnabled;
  var smsPanel = app.createVerticalPanel().setHeight("300px");
  smsPanel.add(app.createLabel("SMS merge").setStyleAttribute("fontSize","16px").setStyleAttribute("borderTop", "1px solid #EBEBEB").setStyleAttribute("borderBottom", "1px solid #EBEBEB").setStyleAttribute("padding","5px").setWidth("100%").setStyleAttribute("backgroundColor", "whiteSmoke"));
  var smsVerticalPanel = app.createVerticalPanel().setWidth("100%").setHeight("280px");
  var smsScrollPanel = app.createScrollPanel().setWidth("560px").setHeight("300px");
  if (smsStatus=="true") {
  if (smsData.length<6) {
    var smsPreviewNum = smsData.length;
  } else {
    var smsPreviewNum = 6
  }
  if (smsPreviewNum>0) {
  for (var i=0; i<smsPreviewNum; i++) {
    var smsLabel = app.createLabel("SMS #"+(i+1)+" - "+smsData[i].num+" Template").setWidth("560px").setStyleAttribute('backgroundColor', '#E5E5E5').setStyleAttribute('textAlign', 'center');
    var grid5 = app.createGrid(2, 2).setBorderWidth(0).setCellSpacing(4).setWidth("100%");
    grid5.setWidget(0, 0, app.createLabel('Phone Number(s):')).setStyleAttribute(0, 0, "width", "100px").setStyleAttribute(0,0,'backgroundColor', 'whiteSmoke');
    grid5.setWidget(1, 0, app.createLabel('Body: (|| shows separation into multiple texts)')).setStyleAttribute(1,0,'backgroundColor', 'whiteSmoke');
    grid5.setWidget(0, 1, app.createLabel(smsData[i].phoneNumber)).setStyleAttribute(0,1,'backgroundColor', 'whiteSmoke');
    grid5.setWidget(1, 1, app.createLabel(smsData[i].body)).setStyleAttribute(1,1,'backgroundColor', 'whiteSmoke');
    smsVerticalPanel.add(smsLabel);
    smsVerticalPanel.add(grid5);
    }
  
    var label5 = app.createLabel('Check out the first ' + smsPreviewNum + ' of the '+ smsData.length +' SMS messages that will be sent.').setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px');
    smsVerticalPanel.add(label5);
   } else {
    var label5 = app.createLabel('No SMS messages will be sent given your current settings.  This could be because of your SMS merge conditions (see Step 2c) or because all your data rows already have a status message in the \'Status\' column for your SMS templates.').setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px');
    smsVerticalPanel.add(label5);
   }
   
  
  if (smsPreviewNum>0) {
    var label6 = app.createLabel("Note: Sender phone number is always that of the Twilio Account.").setStyleAttribute('color', 'blue').setStyleAttribute('marginTop', '2px').setStyleAttribute('marginLeft', '2px');
    smsVerticalPanel.add(label6);
  }
  } else {
    smsVerticalPanel.add(app.createLabel("SMS merge is currently disabled. Want merged SMS Messages? You can enable this service in Step 2c.").setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px'));
  }
  smsScrollPanel.add(smsVerticalPanel);
  smsPanel.add(smsScrollPanel);
 
  tabPanel.add(smsPanel, "SMS");
  
  //End SMS panel
  
   //begin VM panel
  var vmStatus = properties.vmEnabled;
  var vmPanel = app.createVerticalPanel().setHeight("300px");
  vmPanel.add(app.createLabel("VM merge").setStyleAttribute("fontSize","16px").setStyleAttribute("borderTop", "1px solid #EBEBEB").setStyleAttribute("borderBottom", "1px solid #EBEBEB").setStyleAttribute("padding","5px").setWidth("100%").setStyleAttribute("backgroundColor", "whiteSmoke"));
  var vmVerticalPanel = app.createVerticalPanel().setWidth("100%").setHeight("300px");
  var vmScrollPanel = app.createScrollPanel().setWidth("560px").setHeight("300px");
  if (vmStatus=="true") {
  if (vmData.length<6) {
    var vmPreviewNum = vmData.length;
  } else {
    var vmPreviewNum = 6
  }
  if (vmPreviewNum>0) {
    var recordingSheet = ss.getSheetByName("MyVoiceRecordings");
    if (recordingSheet.getLastRow()-1>0) {
      var recordingSelectTexts = recordingSheet.getRange(2, 1, recordingSheet.getLastRow()-1, 1).getValues();
      var recordingSelectValues = recordingSheet.getRange(2, 2, recordingSheet.getLastRow()-1, 1).getValues();
    } else {
      var recordingSelectTexts = [['No messages recorded yet']];
      var recordingSelectValues = [['']];
    }
   
  for (var i=0; i<vmPreviewNum; i++) {
    var vmLabel = app.createLabel("VM #"+(i+1)+" - "+vmData[i].num+" Template").setWidth("560px").setStyleAttribute('backgroundColor', '#E5E5E5').setStyleAttribute('textAlign', 'center');
    var grid6 = app.createGrid(4, 2).setBorderWidth(0).setCellSpacing(4).setWidth("100%");
    grid6.setWidget(0, 0, app.createLabel('Phone Number(s):')).setStyleAttribute(0, 0, "width", "100px").setStyleAttribute(0,0,'backgroundColor', 'whiteSmoke');
    grid6.setWidget(1, 0, app.createLabel('RoboCall Body:')).setStyleAttribute(1,0,'backgroundColor', 'whiteSmoke');
    grid6.setWidget(2, 0, app.createLabel('Play recording:')).setStyleAttribute(2,0,'backgroundColor', 'whiteSmoke');
    grid6.setWidget(3, 0, app.createLabel('Record reply?')).setStyleAttribute(3,0,'backgroundColor', 'whiteSmoke');
    grid6.setWidget(0, 1, app.createLabel(vmData[i].phoneNumber)).setStyleAttribute(0,1,'backgroundColor', 'whiteSmoke');
    grid6.setWidget(1, 1, app.createLabel(vmData[i].body)).setStyleAttribute(1,1,'backgroundColor', 'whiteSmoke');
    var soundFileIndex = -1;
    for (var k=0; k<recordingSelectValues.length; k++) {
      if (recordingSelectValues[k][0]==vmData[i].SoundFile) {
        soundFileIndex = k;
      }
    }
    
    if ((soundFileIndex==-1)&&(!vmData[i].SoundFile||(vmData[i].SoundFile != ''))) {
      var soundFileName = "Not Found";
      grid6.setWidget(2, 1, app.createLabel(soundFileName)).setStyleAttribute(2,1,'backgroundColor', 'whiteSmoke');
    } 
    else if (soundFileIndex!=-1) {
      var soundFileName = recordingSelectTexts[soundFileIndex];
      grid6.setWidget(2, 1, app.createAnchor(soundFileName, vmData[i].SoundFile)).setStyleAttribute(2,1,'backgroundColor', 'whiteSmoke'); 
    }
    if ((!vmData[i].SoundFile)||(vmData[i].SoundFile == '')) {
      var soundFileName = "None";
      grid6.setWidget(2, 1, app.createLabel(soundFileName)).setStyleAttribute(2,1,'backgroundColor', 'whiteSmoke');
    }

    grid6.setWidget(3, 1, app.createLabel(vmData[i].RequestResponse)).setStyleAttribute(3,1,'backgroundColor', 'whiteSmoke');
    vmVerticalPanel.add(vmLabel);
    vmVerticalPanel.add(grid6);
    }
    
    var label7 = app.createLabel('Check out the first ' + vmPreviewNum + ' of the '+ vmData.length +' Voice messages that will be sent.').setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px');
    vmVerticalPanel.add(label7);
   } else {
    var label8 = app.createLabel('No Voice messages will be sent given your current settings.  This could be because of your Voice merge conditions (see Step 2c) or because all your data rows already have a status message in the \'Status\' column for your VM templates.').setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px');
    vmVerticalPanel.add(label8);
   }
    
 
  
  if (vmPreviewNum>0) {
    var label9 = app.createLabel("Note: Caller phone number is always that of the Twilio Account.").setStyleAttribute('color', 'blue');
    vmVerticalPanel.add(label9);
  }
  } else {
    vmVerticalPanel.add(app.createLabel("Voice merge is currently disabled. Want merged Voice Messages? You can enable this service in Step 2c.").setStyleAttribute('fontSize', '14px').setStyleAttribute('marginTop', '5px').setStyleAttribute('marginLeft', '5px'));
  }
  vmScrollPanel.add(vmVerticalPanel);
  vmPanel.add(vmScrollPanel);
  tabPanel.add(vmPanel, "Voice");
  
  //End VM panel

  var runHandler = app.createServerHandler("formMule_manualSend");
  var exitHandler = app.createServerHandler("formMule_quitUi");
  var buttonPanel = app.createHorizontalPanel().setStyleAttribute('marginTop', '10px');
  var button1 = app.createButton("Exit", exitHandler);
  var button2 = app.createButton("Run merge now", runHandler).addClickHandler(exitHandler).setEnabled(false);
  buttonPanel.add(button1).add(button2);
  if ((previewNum>0)||(previewCalNum>0)||(previewUpdateCalNum>0)||(smsPreviewNum>0)||(vmPreviewNum>0)) {
    button2.setEnabled(true);
  }
  tabPanel.selectTab(0);
  panel.add(tabPanel);
  panel.add(buttonPanel);
  app.add(panel);
  ss.show(app);
  return app;
} // end function
//define some useful globals
var https = "https://"
var appName = "api.twilio.com/";
var twilioVersion = "2010-04-01/";
var accounts = "Accounts/"
var webName = appName + twilioVersion + accounts;
var googleLangList = ['Afrikaans: af','Albanian: sq','Arabic: ar','Azerbaijani: az','Basque: eu','Bengali: bn','Belarusian: be','Bulgarian: bg','Catalan: ca','Chinese Simplified: zh-CN','Chinese Traditional: zh-TW','Croatian: hr','Czech: cs','Danish: da','Dutch: nl','English: en','Esperanto: eo','Estonian: et','Filipino: tl','Finnish: fi','French: fr','Galician: gl','Georgian: ka','German: de','Greek: el','Gujarati: gu','Haitian Creole: ht','Hebrew: iw','Hindi: hi','Hungarian: hu','Icelandic: is','Indonesian: id','Irish: ga','Italian: it','Japanese: ja','Kannada: kn','Korean: ko','Latin: la','Latvian: lv','Lithuanian: lt','Macedonian: mk','Malay: ms','Maltese: mt','Norwegian: no','Persian: fa','Polish: pl','Portuguese: pt','Romanian: ro','Russian: ru','Serbian: sr','Slovak: sk','Slovenian: sl','Spanish: es','Swahili: sw','Swedish: sv','Tamil: ta','Telugu: te','Thai: th','Turkish: tr','Ukrainian: uk','Urdu: ur','Vietnamese: vi','Welsh: cy','Yiddish: yi'];
var langList = ['en','en-gb','es','fr','de'];



//function used during installation to verify that the user has in fact made their webapp usable by anonymous visitors (i.e. Twilio)
function formMule_verifyPublic() {
  var url = ScriptApp.getService().getUrl();
  var test = "FALSE";
  if (!url) {
    test = "Error: You have not published your script as a web app. Please revisit the SMS and Voice instructions.";
    return test;
  }
  url +=  "?Verify=TRUE"
  var content = UrlFetchApp.fetch(url).getContentText();
  try {
    test = Xml.parse(content);
    test = test.verified.getText();
    return test;
  } catch(err) {
    test = 'Error: You have published your script as a web app, but you have not made it available to "Anyone, even anonymous."  Please revisit the SMS and Voice instructions.';
    return test;
  }
}


// Ui providing step by step instructions for how to publish the script as a webApp
function formMule_howToPublishAsWebApp() {
  var app = UiApp.createApplication().setTitle('Step 2c. Set up SMS and Voice - Publish formMule as a web app').setHeight(540).setWidth(600);
  var thisSs = SpreadsheetApp.getActiveSpreadsheet();
  ScriptProperties.setProperty('ssId', thisSs.getId());
  var public = formMule_verifyPublic();
  // Once it is determined that the webapp is public, proceed to Twilio setup instructions
  if (public=="TRUE") {
    app.close();
    formMule_howToSetUpTwilio();
    return app;
  }
  var panel = app.createVerticalPanel();
  var handler = app.createServerHandler('formMule_howToSetUpTwilio').addCallbackElement(panel);
  var button = app.createButton("Confirm my settings").addClickHandler(handler);
  var scrollpanel = app.createScrollPanel().setHeight("360px");
  var grid = app.createGrid(9, 2).setBorderWidth(0).setCellSpacing(0);
  var html = app.createHTML('<strong>Important to understand:</strong> Using formMule with the Twilio service is OPTIONAL, but requires publishing this script as a web app, which will provide a URL that Twilio will use to communicate with the script. The instructions below explain how to publish your form as a web app.');
  panel.add(html);
  var text1 = app.createLabel("Instructions:").setStyleAttribute("width", "100%").setStyleAttribute("backgroundColor", "grey").setStyleAttribute("color", "white").setStyleAttribute("padding", "5px 5px 5px 5px");
  panel.add(text1);
  grid.setWidget(0, 0, app.createHTML('1. Go to \'Tools->Script editor\' from the Spreadsheet that contains your form.</li>'));
  grid.setWidget(0, 1, app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/images-for-formmutant/tools%20menu.png?attachauth=ANoY7crn7wMOEG9WpO7uZVHw6b6t5H2WtOs3S_lZ7qV8i2uQsizc1DtFE0HUGG3zXDbqa65wbz3YVJ_9m9zGr-wujYf0jRgz-U45s9Sids1HmI02fPwWNCjkp1GDmzViBwcuL-MwlPNu_WOz4J282Sv1EdVOFQ9l2dKrhPwwOPko1QcWqJ8-v2hS1jOxJaM6VC7EbybBeeLJ1Fy7Z9o56lg1fGcqc28WRklYD3pb_zSh_q_WLfwnl-l0lVQ8HjiWzGNyE5kn4J0Y&attredirects=0').setWidth("430px"));
  grid.setStyleAttribute(1, 0, "backgroundColor", "grey").setStyleAttribute(1, 1, "backgroundColor", "grey");
  grid.setWidget(2, 0, app.createHTML('2. Under the \'File\' menu in the Script Editor, select \'Manage versions\' and save a new version of the script. Because it\'s optional, you can leave the \'Describe what changed\' field blank.</li>'));
  grid.setWidget(2, 1, app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/images-for-formmutant/file%20menu.png?attachauth=ANoY7cqYPycZqgCH7TseuZRL9wq2lv7xI2bhXhDuQ1BVMp5VkB83wpbQRz3X2q36ouPGZyds-CPzEdX5_o-RfDJfVruf7TUxZddxm9sQQwlY6sUab9nVn8KbDgiqiED0itUzfWKvMfP86ya0twRDnONOtGQmq1XWw26vFLQ82DeaKLASttZ6Mo_Qg-FBloN0iVsMlmLLd5o40C4RbaK7qkXUp5eoEE3eYPf_dpx7u3gBb6gyjIZ8yCSKbajC-HfG32MeS0UtrxTW&attredirects=0').setWidth("430px"));
  grid.setWidget(3, 1, app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/images-for-formmutant/manage%20versions.png?attachauth=ANoY7cpVfU-WEwiPuOMKMIAzbK6EdA8xkmv_M2R8GKdlcGLC7mo00ZJykbBFrtJEZQHDpKVdvizQQnuyfGVc65iigmGuGr_ZwC2Z4rnh1V67_ogOJKXH2TWmDAafxa-q_5fngrasDYYN2w2-hR_eR95GoY6e5Rza-mtWb1iAp97Cm8n9kVHRk67dURdrdD5AIaS8ZOkse1MmfaN-ZJpMv7bLYBKpisq8GldTTjo7W55OUIJhFuDcxLEc__vguXArjfb9Pd_e2bZD&attredirects=0').setWidth("430px"));
  grid.setStyleAttribute(4, 0, "backgroundColor", "grey").setStyleAttribute(4, 1, "backgroundColor", "grey");
  grid.setWidget(5, 0, app.createHTML('3. Under the \'Publish\' menu in the Script Editor, select \'Deploy as web app\'. Choose the version you want to publish (usually #1). Under \'Execute web app as\', choose \'me\'. For twilio to access the script, the URL must be visible to anonymous users, so select \'Anyone, even anonymous\'. In the context of this script, this setting can reveal no data to anonymous users.'));
  grid.setWidget(5, 1, app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/images-for-formmutant/publish.png?attachauth=ANoY7crNAwLK0jZOIWN9M0oD2EzLC5pDflZJNnG9pZQcUD4liYn3lBFe_5yr9Hlnc5UgU837A28Mow5hsWQuIJh-ihgu_s_wtDqG51m697X7aFrKPBJoE8WQZRhzgPhLKEh_gYi4Cs9KXJb4Ie-suQb65WbOCvxC4fyStExxZrAcYYcws0U5MJqoStfNqLbvb5iBgys8G6IMSce3tNo-cad8UKVW18aB7J0BLPeTO7spj0u-cJVFXVFFPq22pKwCViQSciRyNYaF&attredirects=0').setWidth("430px"));
  grid.setWidget(6, 1, app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/images-for-formmutant/webapp.png?attachauth=ANoY7co9bP7Sb4cK3aN9Vf3l57RFiWpQ8fDG6_XbNkx8hG7offLTHQAPXm6JMgHjgpEsruUr5mVnf6ha-pc6phil_1MFWsPs7nD2eMeKV_hylLf0qwXLKS3psU4Mu36iw1xvTG8lA6yUJLHb8jA-yRW1jol3bhBD9LKxTzjavHqjtMffZXohjXIV_d_diuNzK6DYBySmuDzeXIp4LyEabrwPglXk5rN6LOXcJ9vPTunUe5cSLhKlLLW2Vj17A3qwmWXxf94gx-yN&attredirects=0').setWidth("430px"));
  grid.setWidget(7, 1, app.createImage('https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/images-for-formmutant/webapp2.png?attachauth=ANoY7crqHkuxcVazt1Bi1_aJwGchj2iACUhPxWHTCa4c8Qrh_jJSgbQQA1CajpsNvIIllJtUIH3yQxCxH-ZiN0wbk_L03Ow7dJorxbqVU5LUjV1ic7gowIbdKhTqS3b3-Sg81PEppEzUlfA8slzI9goYt1KxAXy86RjNHIKUDV1Zo3nORWV7_18GJ5DMpiNB84JV8yHer041WB18JmXYZzpSSJwN3ej00GWo9hknfayVCNy9EhM-QAH8q7pUYJyzwdIRAcW2cayC&attredirects=0').setWidth("430px"));
  scrollpanel.add(grid);
  panel.add(scrollpanel);
  panel.add(button);
  app.add(panel);
  thisSs.show(app);
  return app; 
}


// Ui showing step by step instructions for how to set up Twilio
function formMule_howToSetUpTwilio() {
  var url = ScriptApp.getService().getUrl();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var app = UiApp.createApplication().setTitle("Step 2c. Set up SMS and Voice - Get formMule and Twilio talking").setWidth(600).setHeight(500);
  var public = formMule_verifyPublic();
  if (public!="TRUE") {
    app.close();
    Browser.msgBox(public);
    formMule_howToPublishAsWebApp();
    return app;
  }
  var scrollPanel = app.createScrollPanel().setId("scrollPanel").setWidth("595px").setHeight("410px");
  var panel = app.createVerticalPanel().setId("panel");
  var htmlBody = "Congrats on getting formMule published correctly as a web app.  Now on to configuring Twilio, the 3rd party, paid service that can enable SMS and Voice messages to be sent from this script with a few simple setup steps. Twilio's pricing can be found at http://www.twilio.com/voice/pricing<br>";
  var html = app.createHTML(htmlBody);
  app.add(html);
  var linkLabel = app.createLabel("This script's URL for Twilio SMS and Voice Requests:").setStyleAttribute("margin", "10px 0px 10px 0px").setStyleAttribute("padding", "5px 5px 5px 5px");
  var link = app.createTextBox()
  if (!url) {
    Browser.msgBox("Error: You have not yet published this script as a web app. Please launch Step 2c. again and complete the steps outlined in the instructions.");
    app.close();
    return app;
  } else { 
    ScriptProperties.setProperty('webAppUrl', url);
  }
  link.setText(url).setEnabled(false).setWidth("400px").setStyleAttribute("margin", "10px 10px 10px 10px").setStyleAttribute("padding", "5px 5px 5px 5px");
  var text1 = app.createLabel("Instructions:").setStyleAttribute("width", "100%").setStyleAttribute("backgroundColor", "grey").setStyleAttribute("color", "white").setStyleAttribute("padding", "5px 5px 5px 5px");
  app.add(text1);
  var text2 = app.createLabel("1. Create an account at http://www.twilio.com");
  var image2 = app.createImage("https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/2.jpg?attachauth=ANoY7comDBRVlYUnW6rFi__cI_prSJtORGUDtTjJcGG9QpEI7kNh2MW-u8sxwuB8Ndvs9Nv2HbKQeN2B83DC3YUMjQW-fVd95aMTI7tChp_lxGwAszP5y4ylN5A4rf2_GQhN5nVxv9NdCVf5SfPN1hQiTIS1lNvjRlzjB5voEhbBXq9JLK_apapm6ZHn6qFWfd1vweN6TCtZPiORGJpz235vvxPDxliNZCOvjrzCjf_i-qATTn0TSq4iYy-hHbuUQO2Wpo8zNgtg&attredirects=0").setWidth("430px");
  var text3 = app.createLabel("2. Verify yourself as a human (Twilio will call you and ask you to verify a code), get assigned a free phone number, and proceed to the \"Numbers\" section of your account.  Click on your Twilio phone number.");
  var image3 = app.createImage("https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/6.jpg?attachauth=ANoY7cpOcS5hG8YQwcs01G5LZYz417BhmzEa3lmQgkp-sFSS-xEdE5mu0kFkuuiQ0rwsE6H18OhnqfrXCODtG6FD0OHp8wXynSdaC6kgDb8TCH0OZD55QEgr7Uc3ltdWmj7RkP9BFaOsJ6cxtFMDtpzD6gvJ0CJTr6St7HvdAtTAvTV3XHi75HUlYLCl2drEaAKvDZr396m04Hv_Ohb2ueaH2FtNxaFBZI07SqsPu6pgQC63bIAiCTPEOBq88nCLL0xqCab_mLVZ&attredirects=0").setWidth("430px");
  var numberLabel = app.createLabel("Enter your Twilio number here");
  var numberBox = app.createTextBox().setName("twilioNumber");
  var text4 = app.createLabel("3. Paste the URL of your published script (below) into the Voice and SMS Request URL fields.  Change both URL methods to GET.  Save Changes");
  var image4 = app.createImage("https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/9.jpg?attachauth=ANoY7cp8HufXi1K_3FamZ8SbLFEd3_9DG0ToNsR7vYTvrIzHyKieNlJZn42NJamy9NKxyXdDLgRIl9hXPRx3oWxVRJSVVss3TrFVw2Q2HM28XmylowPDOFyYNLjDW1r-acFaoag74hxzKSLyUAQ6CUB5J_Zheof1Q6XLuE1JZf5dTTXYOtvPANMvrBWrRjOwJ2u9tUnNXAcJfb3BG11MClNWV4d6t1woiexK00zaU6l9bi3QgKDRFZEJXzNhYDnvvlP6qA_SaZQr&attredirects=0").setWidth("430px");
  var fieldPanel = app.createGrid(1,4);
  var accountLabel = app.createLabel("Account SID");
  var accountLabelBox = app.createTextBox().setName('accountSID');
  var accountSID = ScriptProperties.getProperty('accountSID');
  if (accountSID) {
    accountLabelBox.setText(accountSID);
  }
  var authTokenLabel = app.createLabel("Auth Token");
  var authTokenBox = app.createTextBox().setName('authToken');
  var authToken = ScriptProperties.getProperty('authToken');
  if (authToken) {
    authTokenBox.setText(authToken);
  }
  var text5 = app.createLabel("4. Paste and save your Account SID and Auth Token below");
  fieldPanel.setWidget(0, 0, accountLabel).setWidget(0, 1, accountLabelBox).setWidget(0, 2, authTokenLabel).setWidget(0, 3, authTokenBox);
  var image5 = app.createImage("https://c04a7a5e-a-3ab37ab8-s-sites.googlegroups.com/a/newvisions.org/data-dashboard/searchable-docs-collection/1.jpg?attachauth=ANoY7cqnvcVDrWwnN9p3T_tZ1fJ2eDb-AS4JszGViXjYExOVyaW8u6m4CdUGGw460rm7eYhgbKZWqwO4DM9ugkI7vTWX23ZCFVPBE7XiPyz8l6Oeie7To7GGOJKXJTo4Lbs6ufiLVST7iQpgW-0Q3uRKvozd3jUFfrRKO4TMozdZdEjQnLAF2XECiBR1gGBNcxfjJJs1sxAFJNeBuNf5u4jvtt9BYXCD7NGRBRAm2Tp2TiTUNctwbQUqHx8i6-AFYsvSmuM-jfiL&attredirects=0").setWidth("430px");
  var saveHandler = app.createServerHandler('formMule_saveTwilioSettings').addCallbackElement(scrollPanel);
  var button = app.createButton("Verify settings").addClickHandler(saveHandler).setId("testButton");
  var grid = app.createGrid(14, 2).setBorderWidth(0).setCellSpacing(0).setId("grid");
  grid.setWidget(0, 0, text2).setWidget(0, 1, image2);
  grid.setStyleAttribute(1, 0, "backgroundColor", "grey").setStyleAttribute(1, 1, "backgroundColor", "grey");
  grid.setWidget(2, 0, text3).setWidget(2, 1, image3);
  grid.setStyleAttribute(3, 0, "backgroundColor", "grey").setStyleAttribute(3, 1, "backgroundColor", "grey");
  grid.setWidget(4, 0, text4).setWidget(4, 1, image4);
  grid.setWidget(5, 0, linkLabel).setWidget(5, 1, link);
  grid.setStyleAttribute(6, 0, "backgroundColor", "grey").setStyleAttribute(6, 1, "backgroundColor", "grey");
  grid.setWidget(7, 0, text5).setWidget(7, 1, image5);
  grid.setWidget(8, 0, button).setWidget(8, 1, fieldPanel);
  scrollPanel.add(grid);
  app.add(scrollPanel);
  ss.show(app);
}


//Saves Twilio settings
function formMule_saveTwilioSettings(e) {
  var app = UiApp.getActiveApplication();
  var testButton = app.getElementById('testButton');
  var accountSID = e.parameter.accountSID;
  var authToken = e.parameter.authToken;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();
  ScriptProperties.setProperty("ssId", ssId);
  if ((accountSID!='')&&(authToken!='')) {
    ScriptProperties.setProperty("accountSID", accountSID);
    ScriptProperties.setProperty("authToken", authToken);
  }
  var grid = app.getElementById("grid");
  var numberList  = app.createListBox().setName("twilioNumber");
  var phoneNumbers = formMule_getPhoneNumbers(accountSID,authToken);
  var saveHandler = app.createServerHandler('formMule_saveTwilioNumber').addCallbackElement(grid);
  var button = app.createButton("Save settings").setEnabled(false).addClickHandler(saveHandler);
  for (var i = 0; i<phoneNumbers.length; i++) {
    numberList.addItem(phoneNumbers[i]);
  }
  grid.setStyleAttribute(9, 0, "backgroundColor", "grey").setStyleAttribute(9, 1, "backgroundColor", "grey");
  // If no phone numbers are returned by Twilio, warn the user
  if ((!phoneNumbers)||(phoneNumbers.length==0)) {
    button.setEnabled(false);
    testButton.setEnabled(true);   
    grid.setWidget(10, 0, app.createLabel("There appears to be a problem with your settings. Please double check all steps and try again."))
  } else {
    button.setEnabled(true);
    testButton.setEnabled(false);
    grid.setWidget(10, 0, app.createLabel("Twilio phone number(s) successfully detected. If more than one, select which one you want this script to use."));
    var twilioNumber = ScriptProperties.getProperty('twilioNumber');
    if ((twilioNumber)&&(phoneNumbers.indexOf(twilioNumber)!=-1)) {
      numberList.setSelectedIndex(phoneNumbers.indexOf(twilioNumber));
    }
    grid.setWidget(10, 1, numberList);
    grid.setStyleAttribute(11, 0, "backgroundColor", "grey").setStyleAttribute(11, 1, "backgroundColor", "grey");
    var langLabel = app.createLabel("Select default language (from those supported by Twilio) for voice messages");
    var langSelectBox = app.createListBox().setName("defaultLang");
    for (var i=0; i<langList.length; i++) {
      langSelectBox.addItem(langList[i]);
    }
    var langSelected = ScriptProperties.getProperty('defaultVoiceLang');
    if (langSelected) {
      var index = langList.indexOf(langSelected);
      if (index!=-1) {
        langSelectBox.setSelectedIndex(index);
      }
    }
    grid.setWidget(12, 0, langLabel);
    grid.setWidget(12, 1, langSelectBox);
    grid.setWidget(13, 0, button);
  }
  return app;
}

function formMule_saveTwilioNumber(e) {
  var app = UiApp.getActiveApplication();
  var twilioNumber = e.parameter.twilioNumber;
  ScriptProperties.setProperty('twilioNumber', twilioNumber);
  var defaultLang = e.parameter.defaultLang;
  ScriptProperties.setProperty('defaultVoiceLang', defaultLang);
  onOpen();
  Browser.msgBox("Nice work! Now for a bit of magic: Try texting \n \"Woot woot\" to " + twilioNumber + " to see if the script is able to receive and send text messages");
  app.close();
  return app;
}


function formMule_smsAndVoiceSettings(tabIndex) {
  var properties = ScriptProperties.getProperties();
  if (!tabIndex) { 
    tabIndex = 0;
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = ScriptProperties.getProperty('sheetName');
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.getSheets()[0];
    sheetName = sheet.getName();
    ScriptProperties.setProperty('sheetName', sheetName);
    Browser.msgBox("No source sheet detected. Source was automatically set to the top sheet: " + sheetName);
  }
  try {
    var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  } catch(err) {
    sheet.getRange(1,1,1,3).setValues([['Dummy Header 1', 'Dummy Header 2', 'Dummy Header 3']]);
    sheet.setFrozenRows(1);
    Browser.msgBox("Your source sheet must have headers to complete this step");
    formMule_smsAndVoiceSettings();
    return;
  }
  var app = UiApp.createApplication().setTitle("Step 2c: Set up SMS message and voice message merge").setWidth("640").setHeight("450");
  var mainPanel = app.createVerticalPanel().setId("mainPanel").setWidth("640px").setHeight("400px");
  var refreshPanel = app.createVerticalPanel().setId('refreshPanel').setVisible(false);
  var spinner = app.createImage(MULEICONURL).setWidth(150);
  spinner.setStyleAttribute("position", "absolute");
  spinner.setStyleAttribute("top", "120px");
  spinner.setStyleAttribute("left", "190px");
  spinner.setId("dialogspinner");
  refreshPanel.add(spinner);
  
  
  var smsGrid = app.createGrid(1, 2);
  var vmGrid = app.createGrid(1, 2);
  var tabPanel = app.createTabPanel();
  
  var smsPanel = app.createVerticalPanel().setWidth("430px").setHeight("360px");
  var smsFeaturePanel = app.createVerticalPanel().setStyleAttribute("backgroundColor", "whiteSmoke").setWidth("100%").setHeight("100px");
  var smsEnabledCheckBox = app.createCheckBox("Turn on SMS merge feature").setName("smsEnabled").setStyleAttribute("color", "purple");
  if (properties.smsEnabled=="true") {
    smsEnabledCheckBox.setValue(true);
  }
  var smsOnTriggerCheckBox = app.createCheckBox("Trigger this feature on form submit").setName("smsTrigger").setStyleAttribute("color", "purple");
  if (properties.smsTrigger=="true") {
    smsOnTriggerCheckBox.setValue(true);
  }
  var smsLengthPanel = app.createHorizontalPanel();
  var smsLengthLabel = app.createLabel("Max SMS length").setStyleAttribute("padding", "5px");
  var smsLengthSelect = app.createListBox().setName('smsMaxLength');
  smsLengthSelect.addItem("160 char (1 msg per send)", '1');
  smsLengthSelect.addItem("320 char (2 msgs per send)", '2');
  smsLengthSelect.addItem("480 char (3 msgs per send)", '3');
  if (properties.smsMaxLength) {
      smsLengthSelect.setSelectedIndex(parseInt(properties.smsMaxLength)-1);
    }
  smsLengthPanel.add(smsLengthLabel);
  smsLengthPanel.add(smsLengthSelect);
  var smsNumSelectPanel = app.createHorizontalPanel();
  var smsNumSelectHandler = app.createServerHandler('refreshSmsTemplateGrid').addCallbackElement(tabPanel);
  var smsNumSelect = app.createListBox().setName('smsNumSelected');
  smsNumSelect.addItem('1')
    .addItem('2')
    .addItem('3');
  var smsNumSelected = 1;
  if(properties.smsNumSelected) {
    smsNumSelected = properties.smsNumSelected;
  }
  smsNumSelect.setSelectedIndex(smsNumSelected-1);
  smsNumSelect.addChangeHandler(smsNumSelectHandler);
  var smsNumLabel = app.createLabel('unique possible SMS message(s) per row').setStyleAttribute("padding", "5px");
  smsNumSelectPanel.add(smsNumSelect).add(smsNumLabel);
  
  smsFeaturePanel.setStyleAttribute("padding", "5px");
  smsFeaturePanel.add(smsEnabledCheckBox);
  smsFeaturePanel.add(smsOnTriggerCheckBox);
  smsFeaturePanel.add(smsLengthPanel);
  smsFeaturePanel.add(smsNumSelectPanel);
  smsPanel.add(smsFeaturePanel)
  
  var smsTemplatePanel = app.createVerticalPanel().setHeight("250px")
  var smsTemplateScrollPanel = app.createScrollPanel().setHeight("250px");
  var smsTemplateGrid = app.createGrid(smsNumSelected*4, 5).setId('smsTemplateGrid').setBorderWidth(0).setCellSpacing(0);
  var smsPropertyString = properties.smsPropertyString;
  var smsPropertyObject = new Object();
  if ((smsPropertyString)&&(smsPropertyString!='')) {
    smsPropertyObject = Utilities.jsonParse(smsPropertyString);
  }
  
  for (var i = 0; i<smsNumSelected; i++) {
    smsTemplateGrid.setWidget(0+(i*4), 0, app.createLabel('Template Name'));
    smsTemplateGrid.setWidget(1+(i*4), 0, app.createLabel('Phone #'));
    smsTemplateGrid.setWidget(2+(i*4), 0, app.createLabel('Body'));
    smsTemplateGrid.setStyleAttribute(3+(i*4), 0, 'backgroundColor','grey'); 
    var smsTemplateNameBox = app.createTextBox().setName('smsName-'+i).setWidth("140px");
    if (smsPropertyObject['smsName-'+i]) {
      smsTemplateNameBox.setValue(smsPropertyObject['smsName-'+i]);
    }
    var smsPhoneBox = app.createTextBox().setName('smsPhone-'+i).setWidth("140px");
    if (smsPropertyObject['smsPhone-'+i]) {
      smsPhoneBox.setValue(smsPropertyObject['smsPhone-'+i]);
    }
    var smsBodyArea = app.createTextArea().setName('smsBody-'+i).setWidth("140px").setHeight("110px");
    if (smsPropertyObject['smsBody-'+i]) {
      smsBodyArea.setValue(smsPropertyObject['smsBody-'+i]);
    }
    smsTemplateGrid.setWidget(0+(i*4), 1, smsTemplateNameBox);
    smsTemplateGrid.setWidget(1+(i*4), 1, smsPhoneBox);
    smsTemplateGrid.setWidget(2+(i*4), 1, smsBodyArea);
    smsTemplateGrid.setStyleAttribute(3+(i*4), 1, 'backgroundColor','grey'); 
    smsTemplateGrid.setWidget(0+(i*4), 2, app.createLabel('send if'));
    var smsCondCol = app.createListBox().setName('smsCol-'+i).setWidth("90px");
    for (var j=0; j<headers.length; j++) {
      smsCondCol.addItem(headers[j]);
    }
    if (smsPropertyObject['smsCol-'+i]) {
      var index = headers.indexOf(smsPropertyObject['smsCol-'+i]);
      smsCondCol.setSelectedIndex(index);
    }
    smsTemplateGrid.setWidget(1+(i*4), 2, smsCondCol);
    smsTemplateGrid.setWidget(2+(i*4), 2, app.createLabel('Language code (merge tags accepted)').setStyleAttribute('textAlign', 'right'));
    smsTemplateGrid.setStyleAttribute(3+(i*4), 2, 'backgroundColor','grey'); 
    smsTemplateGrid.setWidget(1+(i*4), 3, app.createLabel("="));
    smsTemplateGrid.setStyleAttribute(3+(i*4), 3, 'backgroundColor','grey'); 
    var smsCondVal = app.createTextBox().setName('smsVal-'+i).setWidth("100px");
    if (smsPropertyObject['smsVal-'+i]) {
      smsCondVal.setValue(smsPropertyObject['smsVal-'+i]);
    }
    var smsLang = app.createTextBox().setName('smsLang-'+i).setWidth("50px");
    if (smsPropertyObject['smsLang-'+i]) {
      smsLang.setValue(smsPropertyObject['smsLang-'+i]);
    } else {
      smsLang.setValue('en');
    }
    smsTemplateGrid.setWidget(1+(i*4), 4, smsCondVal);
    smsTemplateGrid.setWidget(2+(i*4), 4, smsLang);
    smsTemplateGrid.setStyleAttribute(3+(i*4), 4, 'backgroundColor','grey'); 
  }
  smsTemplateScrollPanel.add(smsTemplateGrid);
  smsTemplatePanel.add(app.createLabel("SMS Template(s)").setStyleAttribute("color", "white").setStyleAttribute("backgroundColor", "grey").setStyleAttribute("width", "100%").setStyleAttribute('padding', '5px'));
  smsTemplatePanel.add(smsTemplateScrollPanel);
  smsPanel.add(smsTemplatePanel);
  smsGrid.setWidget(0, 0, smsPanel);
  
  
  
  var vmPanel = app.createVerticalPanel().setWidth("430px").setHeight("350px");
  var vmFeaturePanel = app.createVerticalPanel().setStyleAttribute("backgroundColor", "whiteSmoke").setWidth("100%").setHeight("80px").setStyleAttribute("padding", "5px");
  var vmEnabledCheckBox = app.createCheckBox("Turn on Voice Message merge feature").setName('vmEnabled').setStyleAttribute("color", "#FF6600");
  if (properties.vmEnabled=="true") {
    vmEnabledCheckBox.setValue(true);
  }
  var vmOnTriggerCheckBox = app.createCheckBox("Trigger this feature on form submit").setName('vmTrigger').setStyleAttribute("color", "#FF6600");
  if (properties.vmTrigger=="true") {
    vmOnTriggerCheckBox.setValue(true);
  }
  var vmNumSelectPanel = app.createHorizontalPanel();
  var vmNumSelectHandler = app.createServerHandler('refreshVmTemplateGrid').addCallbackElement(tabPanel);
  var vmNumSelect = app.createListBox().setName('vmNumSelected');
  vmNumSelect.addItem('1')
    .addItem('2')
    .addItem('3');
  var vmNumSelected = 1;
  if(properties.vmNumSelected) {
    vmNumSelected = properties.vmNumSelected;
  }
  vmNumSelect.setSelectedIndex(vmNumSelected-1);
  vmNumSelect.addChangeHandler(vmNumSelectHandler);
  var vmNumLabel = app.createLabel('unique possible Voice Message(s) per row').setStyleAttribute("padding", "5px");
  vmNumSelectPanel.add(vmNumSelect).add(vmNumLabel);
  
  vmFeaturePanel.add(vmEnabledCheckBox);
  vmFeaturePanel.add(vmOnTriggerCheckBox);
  vmFeaturePanel.add(vmNumSelectPanel);
  vmPanel.add(vmFeaturePanel);

  var vmTemplatePanel = app.createVerticalPanel().setHeight("270px")
  var vmTemplateScrollPanel = app.createScrollPanel().setHeight("270px");
  var vmTemplateGrid = app.createGrid(vmNumSelected*6, 5).setId("vmTemplateGrid").setBorderWidth(0).setCellSpacing(0);
  var recordingSheet = ss.getSheetByName("MyVoiceRecordings");
  if (!recordingSheet) {
    recordingSheet = ss.insertSheet("MyVoiceRecordings");
    var recordingHeaders = [['Recording Name','Recording URL']];
    recordingSheet.getRange(1, 1, 1, 2).setValues(recordingHeaders).setComment("Don't change the name of this column or sheet. formMule needs these to run SMS and Voice merge services.");
    recordingSheet.setColumnWidth(2, 800);
    recordingSheet.setFrozenRows(1);
    Browser.msgBox("formMule just automatically created a sheet to store the URLs of your outbound voice messages");
    formMule_smsAndVoiceSettings();
    app.close();
    return app;
  }
  if (recordingSheet.getLastRow()-1>0) {
    var recordingSelectTexts = recordingSheet.getRange(2, 1, recordingSheet.getLastRow()-1, 1).getValues();
    var recordingSelectValues = recordingSheet.getRange(2, 2, recordingSheet.getLastRow()-1, 1).getValues();
  } else {
     var recordingSelectTexts = [];
    var recordingSelectValues = [];
  }
  
  var vmPropertyString = properties.vmPropertyString;
  var vmPropertyObject = new Object();
  if ((vmPropertyString)&&(vmPropertyString!='')) {
    vmPropertyObject = Utilities.jsonParse(vmPropertyString);
  }
  
  for (var i = 0; i<vmNumSelected; i++) {
    vmTemplateGrid.setWidget(0+(i*6), 0, app.createLabel('Template Name'));
    vmTemplateGrid.setWidget(1+(i*6), 0, app.createLabel('Phone #'));
    vmTemplateGrid.setWidget(2+(i*6), 0, app.createLabel('Read Message (leave blank for none)'));
    vmTemplateGrid.setWidget(3+(i*6), 0, app.createLabel('Play recording'))
    vmTemplateGrid.setWidget(4+(i*6), 0, app.createLabel('Record response?'));
    vmTemplateGrid.setStyleAttribute(5+(i*6), 0, 'backgroundColor','grey'); 
    var vmTemplateNameBox = app.createTextBox().setName('vmName-'+i).setWidth("140px");
    if (vmPropertyObject['vmName-'+i]) {
      vmTemplateNameBox.setValue(vmPropertyObject['vmName-'+i]);
    }
    var vmPhoneBox = app.createTextBox().setName('vmPhone-'+i).setWidth("140px");
    if (vmPropertyObject['vmPhone-'+i]) {
      vmPhoneBox.setValue(vmPropertyObject['vmPhone-'+i]);
    }
    var vmBodyArea = app.createTextArea().setName('vmBody-'+i).setWidth("140px").setHeight("90px");
    if (vmPropertyObject['vmBody-'+i]) {
      vmBodyArea.setValue(vmPropertyObject['vmBody-'+i]);
    }
    var vmSoundFileListBox = app.createListBox().setName('vmSoundFile-'+i).setWidth("140px");
    var soundFileArray = [];
    vmSoundFileListBox.addItem("None", "");
    var m = 0;
    for (var k=0; k<recordingSelectTexts.length; k++) {
      vmSoundFileListBox.addItem(recordingSelectTexts[k][0], recordingSelectValues[k][0]);
      soundFileArray.push(recordingSelectValues[k][0]);
    }
    if (vmPropertyObject['vmSoundFile-'+i]) {
      var index = soundFileArray.indexOf(vmPropertyObject['vmSoundFile-'+i]);
      if (index!=-1) {
        vmSoundFileListBox.setSelectedIndex(index+1);
        m = index;
      }
    }
    var vmSoundFileLinkPanel = app.createHorizontalPanel();
    var vmSoundFilePlayLinkRefreshHandler = app.createServerHandler('refreshVmSoundFilePlayLink').addCallbackElement(vmSoundFileListBox).addCallbackElement(vmSoundFileLinkPanel);  
    vmSoundFileListBox.addChangeHandler(vmSoundFilePlayLinkRefreshHandler);
    var linkText = "";
    if ((vmPropertyObject['vmSoundFile-'+i]!="")&&(recordingSelectValues.length>0)) {
      linkText = "Play";
      var vmSoundFilePlayLink = app.createAnchor(linkText,recordingSelectValues[m][0]).setId('vmSoundFileLink-'+i);
    } else {
      var vmSoundFilePlayLink = app.createAnchor(linkText,'').setId('vmSoundFileLink-'+i);
    }
    var vmSoundFilePlayLinkClientHandler = app.createClientHandler().forTargets(vmSoundFilePlayLink).setStyleAttribute('color','grey');
    vmSoundFileListBox.addChangeHandler(vmSoundFilePlayLinkClientHandler);
    var vmSoundFileIndex = app.createTextBox().setValue(i).setName('vmIndex').setVisible(false);
    vmSoundFileLinkPanel.add(vmSoundFilePlayLink).add(vmSoundFileIndex);
    var vmRecordOption = app.createListBox().setName('vmRecordOption-'+i).addItem("No").addItem("Yes");
    if (vmPropertyObject['vmRecordOption-'+i]) {
      if (vmPropertyObject['vmRecordOption-'+i]=="Yes") {
        vmRecordOption.setSelectedIndex(1);
      }
    }
    var vmRecordHandler = app.createServerHandler('recordVm').addCallbackElement(mainPanel);
    var vmRecordButton = app.createButton('Call me to record voice message').addClickHandler(vmRecordHandler);
    vmTemplateGrid.setWidget(0+(i*6), 1, vmTemplateNameBox);
    vmTemplateGrid.setWidget(1+(i*6), 1, vmPhoneBox);
    vmTemplateGrid.setWidget(2+(i*6), 1, vmBodyArea);
    vmTemplateGrid.setWidget(3+(i*6), 1, vmSoundFileListBox);
    vmTemplateGrid.setWidget(3+(i*6), 2, vmSoundFileLinkPanel);
    vmTemplateGrid.setWidget(3+(i*6), 4, vmRecordButton);
    vmTemplateGrid.setWidget(4+(i*6), 1, vmRecordOption);
    vmTemplateGrid.setStyleAttribute(5+(i*6), 1, 'backgroundColor','grey'); 
    vmTemplateGrid.setWidget(0+(i*6), 2, app.createLabel('call if'));
    var vmCondCol = app.createListBox().setName('vmCol-'+i).setWidth("90px");
    for (var j=0; j<headers.length; j++) {
      vmCondCol.addItem(headers[j]);
    }
    if (vmPropertyObject['vmCol-'+i]) {
      var index = headers.indexOf(vmPropertyObject['vmCol-'+i]);
      if (index!=-1) {
        vmCondCol.setSelectedIndex(index);
      }
    }
    vmTemplateGrid.setWidget(1+(i*6), 2, vmCondCol);
    vmTemplateGrid.setWidget(2+(i*6), 2, app.createLabel('Language code (merge tags accepted)').setStyleAttribute('textAlign', 'right'));
    vmTemplateGrid.setStyleAttribute(5+(i*6), 2, 'backgroundColor','grey'); 
    vmTemplateGrid.setWidget(1+(i*6), 3, app.createLabel("="));
    vmTemplateGrid.setStyleAttribute(5+(i*6), 3, 'backgroundColor','grey'); 
    var vmCondVal = app.createTextBox().setName('vmVal-'+i).setWidth("100px");
    if (vmPropertyObject['vmVal-'+i]) {
      vmCondVal.setValue(vmPropertyObject['vmVal-'+i]);
    }
    var vmLang = app.createTextBox().setName('vmLang-'+i).setWidth("50px");
    if (vmPropertyObject['vmLang-'+i]) {
        vmLang.setValue(vmPropertyObject['vmLang-'+i]);
    } else {
        vmLang.setValue("en");
    }
    vmTemplateGrid.setWidget(1+(i*6), 4, vmCondVal);
    vmTemplateGrid.setWidget(2+(i*6), 4, vmLang);
    vmTemplateGrid.setStyleAttribute(5+(i*6), 4, 'backgroundColor','grey'); 
  }
  vmTemplateScrollPanel.add(vmTemplateGrid);
  vmTemplatePanel.add(app.createLabel("Voice Message(s)").setStyleAttribute("color", "white").setStyleAttribute("backgroundColor", "grey").setStyleAttribute("width", "100%").setStyleAttribute('padding', '5px'));
  vmTemplatePanel.add(vmTemplateScrollPanel);
  vmPanel.add(vmTemplatePanel);
  vmGrid.setWidget(0,0,vmPanel);
 
  
  var inboundPanel = app.createVerticalPanel().setWidth("620px").setHeight("360px");
  var inboundSMSPanel = app.createVerticalPanel().setStyleAttribute('backgroundColor', 'whiteSmoke').setWidth("425px");
  var inboundSMSLabel = app.createLabel("Inbound SMS Handling").setStyleAttribute('width', '425px').setStyleAttribute('backgroundColor', 'grey').setStyleAttribute('color', 'white').setStyleAttribute('margin', '5px').setStyleAttribute('padding', '5px');
  var inboundTextCheckBox = app.createCheckBox("send this text reply to all inbound SMS messages").setName('smsAutoReplyOption').setStyleAttribute('margin', '5px');
  if (properties.smsAutoReplyOption == "true") {
    inboundTextCheckBox.setValue(true);
  }
  var inboundTextBody = app.createTextArea().setName('smsAutoReplyBody').setWidth("425px").setStyleAttribute('margin', '5px');
  if (properties.smsAutoReplyBody) {
    inboundTextBody.setValue(properties.smsAutoReplyBody);
  }
  var inboundTextForwardPanel = app.createHorizontalPanel().setStyleAttribute('margin', '5px')
  var inboundTextForwardCheckBox = app.createCheckBox('forward all inbound SMS messages to this email address').setName('smsAutoForwardOption').setStyleAttribute('padding', '5px');
  if (properties.smsAutoForwardOption=="true") {
    inboundTextForwardCheckBox.setValue(true);
  }
  var inboundTextForwardEmail = app.createTextBox().setName('smsAutoForwardEmail').setStyleAttribute('margin', '5px').setWidth("130px");
  if (properties.smsAutoForwardEmail) {
    inboundTextForwardEmail.setValue(properties.smsAutoForwardEmail);
  }
  inboundTextForwardPanel.add(inboundTextForwardCheckBox).add(inboundTextForwardEmail);
  inboundSMSPanel.add(inboundSMSLabel).add(inboundTextCheckBox).add(inboundTextBody).add(inboundTextForwardPanel);
  var inboundVoicePanel = app.createVerticalPanel().setStyleAttribute('backgroundColor', 'whiteSmoke').setWidth("425");
  
  var inboundVoiceLabel = app.createLabel("Inbound Call Handling").setStyleAttribute('width', '425px').setStyleAttribute('backgroundColor', 'grey').setStyleAttribute('color', 'white').setStyleAttribute('margin', '5px').setStyleAttribute('marginTop', '8px').setStyleAttribute('padding', '5px');
  var inboundVoiceCheckBox = app.createCheckBox("read this message to all inbound callers").setName('vmAutoReplyReadOption').setStyleAttribute('margin', '5px');
  if (properties.vmAutoReplyReadOption=="true") {
    inboundVoiceCheckBox.setValue(true);
  }
  var inboundVoiceBody = app.createTextArea().setWidth("425px").setName('vmAutoReplyBody').setStyleAttribute('margin', '5px');
  if (properties.vmAutoReplyBody) {
    inboundVoiceBody.setValue(properties.vmAutoReplyBody);
  }
  var inboundVoiceMessagePanel = app.createHorizontalPanel();
  var inboundVoiceMessageCheckBox = app.createCheckBox("play this voice recording").setName('vmAutoReplyPlayOption').setStyleAttribute('margin', '5px');
  if (properties.vmAutoReplyPlayOption=="true") {
    inboundVoiceMessageCheckBox.setValue(true);
  }
  var inboundSoundFileListBox = app.createListBox().setName('vmAutoReplyFile').setWidth("150px");
  var m = 0;
  var inboundSoundFileChangeHandler = app.createServerHandler('refreshInboundPlayLink').addCallbackElement(inboundVoiceMessagePanel);
  inboundSoundFileListBox.addChangeHandler(inboundSoundFileChangeHandler);
  inboundSoundFileListBox.addItem("None", "")
  for (var k=0; k<recordingSelectTexts.length; k++) {
    inboundSoundFileListBox.addItem(recordingSelectTexts[k][0], recordingSelectValues[k][0]);
    if (properties.vmAutoReplyFile) {
      if (properties.vmAutoReplyFile==recordingSelectValues[k][0]) {
        inboundSoundFileListBox.setSelectedIndex(k+1);
        m = k;
      }
    }
  }
  if (tabIndex==2) {
    inboundSoundFileListBox.setSelectedIndex(recordingSelectTexts.length);
  }
  var inboundRecordHandler = app.createServerHandler('inboundRecordVm').addCallbackElement(mainPanel);
  if (recordingSelectValues.length>0) {
    if (properties.vmAutoReplyFile!='') {
      var inboundPlayLink = app.createAnchor("Play", recordingSelectValues[m][0]).setId('inboundPlayLink');
    } else {
      var inboundPlayLink = app.createLabel("").setId('inboundPlayLink');
    }
  } else {
    var inboundPlayLink = app.createLabel("").setId('inboundPlayLink');
  }
  var inboundSoundFileDissapearHandler = app.createClientHandler().forTargets(inboundPlayLink).setStyleAttribute('color','grey');
  inboundSoundFileListBox.addChangeHandler(inboundSoundFileDissapearHandler);
  var inboundRecordButton = app.createButton('Call me to record voice message').addClickHandler(inboundRecordHandler);
  inboundVoiceMessagePanel.add(inboundVoiceMessageCheckBox).add(inboundSoundFileListBox).add(inboundPlayLink).add(inboundRecordButton);
  var inboundVoiceForwardPanel = app.createHorizontalPanel().setStyleAttribute('margin', '5px').setWidth("425px");
  var inboundVoiceForwardCheckBox = app.createCheckBox('forward all inbound calls to this number').setName('vmAutoForwardOption').setStyleAttribute('margin', '5px');
  if (properties.vmAutoForwardOption=="true") {
    inboundVoiceForwardCheckBox.setValue(true);
  }
  var inboundVoiceForwardNumber = app.createTextBox().setName('vmAutoForwardNumber').setStyleAttribute('margin', '5px');
  if (properties.vmAutoForwardNumber) {
    inboundVoiceForwardNumber.setValue(properties.vmAutoForwardNumber);
  }
  inboundVoiceForwardPanel.add(inboundVoiceForwardCheckBox).add(inboundVoiceForwardNumber);
  inboundVoicePanel.add(inboundVoiceLabel).add(inboundVoiceCheckBox).add(inboundVoiceBody).add(inboundVoiceMessagePanel).add(inboundVoiceForwardPanel);
  inboundPanel.add(inboundSMSPanel).add(inboundVoicePanel);
  tabPanel.add(smsGrid, "SMS Merge").add(vmGrid, "Voice Message Merge").add(inboundPanel,"Inbound Traffic Handling").selectTab(tabIndex);
  mainPanel.add(tabPanel);
  var rightSMSPanel = app.createVerticalPanel(); 
  var rightVMPanel = app.createVerticalPanel(); 
  
  var variablesPanel = app.createVerticalPanel().setStyleAttribute('backgroundColor', 'whiteSmoke').setStyleAttribute('padding', '10px').setBorderWidth(2).setStyleAttribute('borderColor', '#92C1F0').setWidth("100%");
  var variablesScrollPanel = app.createScrollPanel().setHeight(130);
  var variablesLabel = app.createLabel().setText("Merge tags: ").setStyleAttribute('fontWeight', 'bold');
  var tags = formMule_getAvailableTags();
  var flexTable = app.createFlexTable().setBorderWidth(0);
  for (var i = 0; i<tags.length; i++) {
    var tag = app.createLabel().setText(tags[i]);
    flexTable.setWidget(i, 0, tag);
  }
  variablesPanel.add(variablesLabel);
  variablesScrollPanel.add(flexTable);
  variablesPanel.add(variablesScrollPanel);
  
  var variablesPanel2 = app.createVerticalPanel().setStyleAttribute('backgroundColor', 'whiteSmoke').setStyleAttribute('padding', '10px').setBorderWidth(2).setStyleAttribute('borderColor', '#92C1F0').setWidth("100%");
  var variablesScrollPanel2 = app.createScrollPanel().setHeight(130);
  var variablesLabel2 = app.createLabel().setText("Merge tags: ").setStyleAttribute('fontWeight', 'bold');
  var tags2 = formMule_getAvailableTags();
  var flexTable2 = app.createFlexTable().setBorderWidth(0);
  for (var i = 0; i<tags2.length; i++) {
    var tag2 = app.createLabel().setText(tags2[i]);
    flexTable2.setWidget(i, 0, tag2);
  }
  variablesPanel2.add(variablesLabel2);
  variablesScrollPanel2.add(flexTable2);
  variablesPanel2.add(variablesScrollPanel2);


  
  var SMSLangPanel = app.createVerticalPanel().setStyleAttribute('backgroundColor', 'whiteSmoke').setStyleAttribute('padding', '10px').setBorderWidth(2).setStyleAttribute('borderColor', '#92C1F0').setWidth("184px").setStyleAttribute('marginTop', '2px');
  var SMSLangScrollPanel = app.createScrollPanel().setHeight(130);
  var SMSLangLabel = app.createLabel().setText("SMS Language Codes").setStyleAttribute('fontWeight', 'bold');
  var SMSLangs = googleLangList;
  var SMSLangFlexTable = app.createFlexTable().setBorderWidth(0);
  for (var i = 0; i<SMSLangs.length; i++) {
    var tag = app.createLabel().setText(SMSLangs[i]);
    SMSLangFlexTable.setWidget(i, 0, tag);
    }
  SMSLangPanel.add(SMSLangLabel);
  SMSLangScrollPanel.add(SMSLangFlexTable);
  SMSLangPanel.add(SMSLangScrollPanel);
  
  var VMLangPanel = app.createVerticalPanel().setStyleAttribute('backgroundColor', 'whiteSmoke').setStyleAttribute('padding', '10px').setBorderWidth(2).setStyleAttribute('borderColor', '#92C1F0').setWidth("184px").setStyleAttribute('marginTop', '2px');
  var VMLangScrollPanel = app.createScrollPanel().setHeight(130);
  var VMLangLabel = app.createLabel().setText("Voice Language Codes").setStyleAttribute('fontWeight', 'bold');
  var VMLangs = langList;
  var VMLangFlexTable = app.createFlexTable().setBorderWidth(0);
  for (var i = 0; i<VMLangs.length; i++) {
    var tag = app.createLabel().setText(VMLangs[i]);
    VMLangFlexTable.setWidget(i, 0, tag);
    }
  VMLangPanel.add(VMLangLabel);
  VMLangScrollPanel.add(VMLangFlexTable);
  VMLangPanel.add(VMLangScrollPanel);
  

  var rightSMSBottomPanel = app.createVerticalPanel();
  var rightVMBottomPanel = app.createVerticalPanel();
  rightVMBottomPanel.add(VMLangPanel);
  rightSMSBottomPanel.add(SMSLangPanel);
  rightSMSPanel.add(variablesPanel2);
  rightSMSPanel.add(rightSMSBottomPanel);
  
  rightVMPanel.add(variablesPanel);
  rightVMPanel.add(rightVMBottomPanel);
  smsGrid.setWidget(0, 1, rightSMSPanel);
  vmGrid.setWidget(0, 1, rightVMPanel);
  var saveHandler = app.createServerHandler('saveSmsAndVoiceSettings').addCallbackElement(mainPanel);
  var saveSpinnerHandler = app.createClientHandler().forTargets(mainPanel).setStyleAttribute('opacity', '0.5').forTargets(refreshPanel).setVisible(true);  
  var button = app.createButton("Save settings").addClickHandler(saveHandler).addClickHandler(saveSpinnerHandler);
  mainPanel.add(button);
  app.add(mainPanel);
  app.add(refreshPanel);
  ss.show(app);
  return app;
}

function refreshInboundPlayLink(e) {
  var app = UiApp.getActiveApplication();
  var link = app.getElementById('inboundPlayLink');
  var newUrl = e.parameter.vmAutoReplyFile;
 if (newUrl != '') {
    link.setText("Play")
    link.setHref(newUrl);
    link.setStyleAttribute('color','blue');
  } else {
    link.setText('');
  }
  return app;
}

function refreshVmSoundFilePlayLink(e) {
  var app = UiApp.getActiveApplication();
  var i = e.parameter.vmIndex;
  var link = app.getElementById('vmSoundFileLink-'+i);
  var newUrl = e.parameter['vmSoundFile-'+i];
  if (newUrl != '') {
    link.setText("Play")
    link.setHref(newUrl);
    link.setStyleAttribute('color','blue');
  } else {
    link.setText('');
  }
  return app;
}


function refreshSmsTemplateGrid(e) {
  var app = UiApp.getActiveApplication();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = ScriptProperties.getProperty('sheetName');
  var sheet = ss.getSheetByName(sheetName);
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var properties = ScriptProperties.getProperties();
  var smsNumSelected = e.parameter.smsNumSelected;
  properties.smsNumSelected = smsNumSelected;
  
  var smsTemplateGrid = app.getElementById('smsTemplateGrid');
  smsTemplateGrid.resize(smsNumSelected*4, 5);
  var smsPropertyString = properties.smsPropertyString;
  var smsPropertyObject = new Object();
  for (var i=0; i<properties.smsNumSelected; i++) {
    smsPropertyObject['smsName-'+i] = e.parameter['smsName-'+i];
    smsPropertyObject['smsPhone-'+i] = e.parameter['smsPhone-'+i];
    smsPropertyObject['smsBody-'+i] = e.parameter['smsBody-'+i];
    smsPropertyObject['smsLang-'+i] = e.parameter['smsLang-'+i];
    smsPropertyObject['smsCol-'+i] = e.parameter['smsCol-'+i];
    smsPropertyObject['smsVal-'+i] = e.parameter['smsVal-'+i];
  } 
  for (var i = 0; i<smsNumSelected; i++) {
    smsTemplateGrid.setWidget(0+(i*4), 0, app.createLabel('Template Name'));
    smsTemplateGrid.setWidget(1+(i*4), 0, app.createLabel('Phone #'));
    smsTemplateGrid.setWidget(2+(i*4), 0, app.createLabel('Body'));
    smsTemplateGrid.setStyleAttribute(3+(i*4), 0, 'backgroundColor','grey'); 
    var smsTemplateNameBox = app.createTextBox().setName('smsName-'+i).setWidth("140px");
    if (smsPropertyObject['smsName-'+i]) {
      smsTemplateNameBox.setValue(smsPropertyObject['smsName-'+i]);
    }
    var smsPhoneBox = app.createTextBox().setName('smsPhone-'+i).setWidth("140px");
    if (smsPropertyObject['smsPhone-'+i]) {
      smsPhoneBox.setValue(smsPropertyObject['smsPhone-'+i]);
    }
    var smsBodyArea = app.createTextArea().setName('smsBody-'+i).setWidth("140px").setHeight("110px");
    if (smsPropertyObject['smsBody-'+i]) {
      smsBodyArea.setValue(smsPropertyObject['smsBody-'+i]);
    }
    smsTemplateGrid.setWidget(0+(i*4), 1, smsTemplateNameBox);
    smsTemplateGrid.setWidget(1+(i*4), 1, smsPhoneBox);
    smsTemplateGrid.setWidget(2+(i*4), 1, smsBodyArea);
    smsTemplateGrid.setStyleAttribute(3+(i*4), 1, 'backgroundColor','grey'); 
    smsTemplateGrid.setWidget(0+(i*4), 2, app.createLabel('send if'));
    var smsCondCol = app.createListBox().setName('smsCol-'+i).setWidth("90px");
    for (var j=0; j<headers.length; j++) {
      smsCondCol.addItem(headers[j]);
    }
    if (smsPropertyObject['smsCol-'+i]) {
      var index = headers.indexOf(smsPropertyObject['smsCol-'+i]);
      smsCondCol.setSelectedIndex(index);
    }
    smsTemplateGrid.setWidget(1+(i*4), 2, smsCondCol);
    smsTemplateGrid.setWidget(2+(i*4), 2, app.createLabel('Language code (merge tags accepted)').setStyleAttribute('textAlign', 'right'));
    smsTemplateGrid.setStyleAttribute(3+(i*4), 2, 'backgroundColor','grey'); 
    smsTemplateGrid.setWidget(1+(i*4), 3, app.createLabel("="));
    smsTemplateGrid.setStyleAttribute(3+(i*4), 3, 'backgroundColor','grey'); 
    var smsCondVal = app.createTextBox().setName('smsVal-'+i).setWidth("100px");
    if (smsPropertyObject['smsVal-'+i]) {
      smsCondVal.setValue(smsPropertyObject['smsVal-'+i]);
    }
    var smsLang = app.createTextBox().setName('smsLang-'+i).setWidth("50px");
    if (smsPropertyObject['smsLang-'+i]) {
      smsLang.setValue(smsPropertyObject['smsLang-'+i]);
    } else {
      smsLang.setValue('en');
    }
    smsTemplateGrid.setWidget(1+(i*4), 4, smsCondVal);
    smsTemplateGrid.setWidget(2+(i*4), 4, smsLang);
    smsTemplateGrid.setStyleAttribute(3+(i*4), 4, 'backgroundColor','grey'); 
  }
  return app;
}


function refreshVmTemplateGrid(e) {
  var app = UiApp.getActiveApplication();
  var mainPanel = app.getElementById("mainPanel");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = ScriptProperties.getProperty('sheetName');
  var sheet = ss.getSheetByName(sheetName);
  var recordingSheet = ss.getSheetByName("MyVoiceRecordings");
  var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
  var recordingSelectTexts = [];
  var recordingSelectValues = [];
  if (recordingSheet.getLastRow()-1>0) {
    recordingSelectTexts = recordingSheet.getRange(2, 1, recordingSheet.getLastRow()-1, 1).getValues();
    recordingSelectValues = recordingSheet.getRange(2, 2, recordingSheet.getLastRow()-1, 1).getValues();
  } 
  var properties = ScriptProperties.getProperties();
  var vmNumSelected = e.parameter.vmNumSelected;
  properties.vmNumSelected = vmNumSelected;
  var vmTemplateGrid = app.getElementById('vmTemplateGrid');
  vmTemplateGrid.resize(vmNumSelected*6, 5);
  var vmPropertyString = properties.vmPropertyString;
  var vmPropertyObject = new Object();
  for (var i=0; i<properties.vmNumSelected; i++) {
    vmPropertyObject['vmName-'+i] = e.parameter['vmName-'+i];
    vmPropertyObject['vmPhone-'+i] = e.parameter['vmPhone-'+i];
    vmPropertyObject['vmBody-'+i] = e.parameter['vmBody-'+i];
    vmPropertyObject['vmSoundFile-'+i] = e.parameter['vmSoundFile-'+i];
    vmPropertyObject['vmRecordOption-'+i] = e.parameter['vmRecordOption-'+i];
    vmPropertyObject['vmLang-'+i] = e.parameter['vmLang-'+i];
    vmPropertyObject['vmCol-'+i] = e.parameter['vmCol-'+i];
    vmPropertyObject['vmVal-'+i] = e.parameter['vmVal-'+i];
  }
  for (var i = 0; i<vmNumSelected; i++) {
    vmTemplateGrid.setWidget(0+(i*6), 0, app.createLabel('Template Name'));
    vmTemplateGrid.setWidget(1+(i*6), 0, app.createLabel('Phone #'));
    vmTemplateGrid.setWidget(2+(i*6), 0, app.createLabel('Read Message'));
    vmTemplateGrid.setWidget(3+(i*6), 0, app.createLabel('Play recording'))
    vmTemplateGrid.setWidget(4+(i*6), 0, app.createLabel('Record response?'));
    vmTemplateGrid.setStyleAttribute(5+(i*6), 0, 'backgroundColor','grey'); 
    var vmTemplateNameBox = app.createTextBox().setName('vmName-'+i).setWidth("140px");
    if (vmPropertyObject['vmName-'+i]) {
      vmTemplateNameBox.setValue(vmPropertyObject['vmName-'+i]);
    }
    var vmPhoneBox = app.createTextBox().setName('vmPhone-'+i).setWidth("140px");
    if (vmPropertyObject['vmPhone-'+i]) {
      vmPhoneBox.setValue(vmPropertyObject['vmPhone-'+i]);
    }
    var vmBodyArea = app.createTextArea().setName('vmBody-'+i).setWidth("140px").setHeight("90px");
    if (vmPropertyObject['vmBody-'+i]) {
      vmBodyArea.setValue(vmPropertyObject['vmBody-'+i]);
    }
    var vmSoundFileListBox = app.createListBox().setName('vmSoundFile-'+i).setWidth("140px");
    var soundFileArray = [];
    vmSoundFileListBox.addItem("None", "");
    for (var k=0; k<recordingSelectTexts.length; k++) {
      vmSoundFileListBox.addItem(recordingSelectTexts[k][0], recordingSelectValues[k][0]);
      soundFileArray.push(recordingSelectValues[k][0]);
    }
    var m = 0;
    var n = -1;
    if (vmPropertyObject['vmSoundFile-'+i]) {
      var index = soundFileArray.indexOf(vmPropertyObject['vmSoundFile-'+i]);
      if (index!=-1) {
        vmSoundFileListBox.setSelectedIndex(index + 1);
        m = index;
        n = 0;
      }
    } 
    var vmSoundFileLinkPanel = app.createHorizontalPanel();
    var vmSoundFilePlayLinkRefreshHandler = app.createServerHandler('refreshVmSoundFilePlayLink').addCallbackElement(vmSoundFileListBox).addCallbackElement(vmSoundFileLinkPanel);  
    vmSoundFileListBox.addChangeHandler(vmSoundFilePlayLinkRefreshHandler);
    var linkText = "";
    if ((n!=-1)&&(vmPropertyObject['vmSoundFileLink-'+i]!='')&&(recordingSelectValues.length>0)) {
      linkText = "Play";
    }
    var vmSoundFilePlayLink = app.createAnchor(linkText,recordingSelectValues[m][0]).setId('vmSoundFileLink-'+i);
    var vmSoundFilePlayLinkClientHandler = app.createClientHandler().forTargets(vmSoundFilePlayLink).setStyleAttribute('color','grey');
    vmSoundFileListBox.addChangeHandler(vmSoundFilePlayLinkClientHandler);
    var vmSoundFileIndex = app.createTextBox().setValue(i).setName('vmIndex').setVisible(false);
    vmSoundFileLinkPanel.add(vmSoundFilePlayLink).add(vmSoundFileIndex);  
    var vmRecordOption = app.createListBox().setName('vmRecordOption-'+i).addItem("No").addItem("Yes");
    if (vmPropertyObject['vmRecordOption-'+i]) {
      if (vmPropertyObject['vmRecordOption-'+i]=="Yes") {
        vmRecordOption.setSelectedIndex(1);
      }
    }
    var vmRecordHandler = app.createServerHandler('recordVm').addCallbackElement(mainPanel);
    var vmRecordButton = app.createButton('Call me to record voice message').addClickHandler(vmRecordHandler);
    vmTemplateGrid.setWidget(0+(i*6), 1, vmTemplateNameBox);
    vmTemplateGrid.setWidget(1+(i*6), 1, vmPhoneBox);
    vmTemplateGrid.setWidget(2+(i*6), 1, vmBodyArea);
    vmTemplateGrid.setWidget(3+(i*6), 1, vmSoundFileListBox);
    vmTemplateGrid.setWidget(3+(i*6), 2, vmSoundFileLinkPanel);
    vmTemplateGrid.setWidget(3+(i*6), 4, vmRecordButton);
    vmTemplateGrid.setWidget(4+(i*6), 1, vmRecordOption);
    vmTemplateGrid.setStyleAttribute(5+(i*6), 1, 'backgroundColor','grey'); 
    vmTemplateGrid.setWidget(0+(i*6), 2, app.createLabel('call if'));
    var vmCondCol = app.createListBox().setName('vmCol-'+i).setWidth("90px");
    for (var j=0; j<headers.length; j++) {
      vmCondCol.addItem(headers[j]);
    }
    if (properties['vmCol-'+i]) {
      var index = headers.indexOf(properties['vmCol-'+i]);
      if (index!=-1) {
        vmCondCol.setSelectedIndex(index);
      }
    }
    vmTemplateGrid.setWidget(1+(i*6), 2, vmCondCol);
    vmTemplateGrid.setWidget(2+(i*6), 2, app.createLabel('Language code (merge tags accepted)').setStyleAttribute('textAlign', 'right'));
    vmTemplateGrid.setStyleAttribute(5+(i*6), 2, 'backgroundColor','grey'); 
    vmTemplateGrid.setWidget(1+(i*6), 3, app.createLabel("="));
    vmTemplateGrid.setStyleAttribute(5+(i*6), 3, 'backgroundColor','grey'); 
    var vmCondVal = app.createTextBox().setName('vmVal-'+i).setWidth("100px");
    if (properties['vmVal-'+i]) {
      vmCondVal.setValue(properties['vmVal-'+i]);
    }
    var vmLang = app.createTextBox().setName('vmLang-'+i).setWidth("50px");
    if (properties['vmLang-'+i]) {
        vmLang.setValue(vmLang);
    } else {
        vmLang.setValue("en");
    }
    vmTemplateGrid.setWidget(1+(i*6), 4, vmCondVal);
    vmTemplateGrid.setWidget(2+(i*6), 4, vmLang);
    vmTemplateGrid.setStyleAttribute(5+(i*6), 4, 'backgroundColor','grey'); 
  }
  return app;
}


function inboundRecordVm(e) {
  var app = UiApp.getActiveApplication();
  app.close();
  e.parameter.num = "2";
  inboundRecordApp(e)
  return app;
}

function recordVm(e) {
  var app = UiApp.getActiveApplication();
  app.close();
  e.parameter.num = "1";
  inboundRecordApp(e)
  return app;
  
}

function inboundRecordApp(e) {
  saveSmsAndVoiceSettings(e);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('MyVoiceRecordings');
  ss.setActiveSheet(sheet);
  var app = UiApp.createApplication().setTitle("Record a new message").setHeight(200);
  var panel = app.createVerticalPanel();
  var phoneNumberLabel = app.createLabel("Please enter the phone number you would like Twilio to call to record your voice message");
  var phoneNumberBox = app.createTextBox().setName('phone').setWidth('150px');
  var lastPhone = ScriptProperties.getProperty('lastPhone');
  if (lastPhone) {
    phoneNumberBox.setValue(lastPhone);
  }
  var recordingName = app.createLabel("Please give a unique name to this recording");
  var recordingNameBox = app.createTextBox().setName('name').setWidth('150px');
  var hiddenNumBox = app.createTextBox().setName('num').setVisible(false).setValue(e.parameter.num);
  var panel2 = app.createVerticalPanel().setId('callingPanel').setVisible(false);
  var callingImage = app.createImage('http://status.twilio.com/images/logo.png').setWidth("100px");
  var callingLabel = app.createLabel('Twilio is attempting to call the number you provided.  This can sometimes take 30-45 seconds. Click \'Done\' once you have hung up and can see your recording show up in the \'MyVoiceRecordings\' sheet');
  var doneButtonHandler = app.createServerHandler('doneRecording').addCallbackElement(panel);
  var doneButton = app.createButton('Done').addClickHandler(doneButtonHandler);
  panel2.add(callingImage).add(callingLabel).add(doneButton);
  panel.add(phoneNumberLabel).add(phoneNumberBox).add(recordingName).add(recordingNameBox).add(hiddenNumBox);
  var buttonClientHandler = app.createClientHandler().forTargets(panel2).setVisible(true).forTargets(panel).setVisible(false);
  var buttonHandler = app.createServerHandler('saveRecording').addCallbackElement(panel);
  var button = app.createButton('Call me to record my message').addClickHandler(buttonHandler).addClickHandler(buttonClientHandler);
  panel.add(button);
  app.add(panel);
  app.add(panel2);
  ss.show(app);
  return app;
}


function saveRecording(e) {
  var app = UiApp.getActiveApplication();
  var num = e.parameter.num;
  var phoneNumber = e.parameter.phone;
  ScriptProperties.setProperty('lastPhone', phoneNumber);
  var recordingName = e.parameter.name;
  var accountSID = ScriptProperties.getProperty("accountSID");
  var authToken = ScriptProperties.getProperty("authToken");
  var twilioNumber = ScriptProperties.getProperty("twilioNumber");
  var args = new Object();
  args.RequestResponse = "TRUE";
  args.OutboundMessageRecording = "TRUE";
  args.RecordingName = recordingName;
  var lang = ScriptProperties.getProperty('defaultVoiceLang');
  formMule_makeRoboCall(phoneNumber, "Please record your message after the tone.", accountSID, authToken, '', args, lang);
  return app;
}

function doneRecording(e) {
  var app = UiApp.getActiveApplication();
  var num = parseInt(e.parameter.num);
  app.close();
  formMule_smsAndVoiceSettings(num);
  return app;
}


function saveSmsAndVoiceSettings(e) {
  var app = UiApp.getActiveApplication();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var properties = ScriptProperties.getProperties();
  var sheet = ss.getSheetByName(properties.sheetName);
  var headers = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues()[0];
  properties.smsEnabled = e.parameter.smsEnabled;
  properties.smsTrigger = e.parameter.smsTrigger;
  properties.smsMaxLength = e.parameter.smsMaxLength;
  properties.smsNumSelected = e.parameter.smsNumSelected;
  var smsPropertyObject = new Object();
  for (var i=0; i<properties.smsNumSelected; i++) {
    smsPropertyObject['smsName-'+i] = e.parameter['smsName-'+i];
    if ((headers.indexOf(smsPropertyObject['smsName-'+i] + " SMS Status") == -1)&&(properties.smsEnabled=="true")) {
      sheet.insertColumnAfter(sheet.getLastColumn());
      sheet.getRange(1, sheet.getLastColumn()+1).setValue(smsPropertyObject['smsName-'+i] + " SMS Status").setBackgroundColor('purple').setFontColor('white').setComment('Don\'t change the name of this column. It is used to log the send status of an SMS Message, whose name must be a match.');
      sheet.setColumnWidth(sheet.getLastColumn(), "400");
    }
    smsPropertyObject['smsPhone-'+i] = e.parameter['smsPhone-'+i];
    smsPropertyObject['smsBody-'+i] = e.parameter['smsBody-'+i];
    smsPropertyObject['smsLang-'+i] = e.parameter['smsLang-'+i];
    smsPropertyObject['smsCol-'+i] = e.parameter['smsCol-'+i];
    smsPropertyObject['smsVal-'+i] = e.parameter['smsVal-'+i];
  }
  var smsPropertyString = Utilities.jsonStringify(smsPropertyObject);
  properties.smsPropertyString = smsPropertyString;
  properties.smsAutoReplyOption = e.parameter.smsAutoReplyOption;
  properties.smsAutoReplyBody = e.parameter.smsAutoReplyBody;
  properties.smsAutoForwardOption = e.parameter.smsAutoForwardOption;
  properties.smsAutoForwardEmail = e.parameter.smsAutoForwardEmail;
  properties.vmAutoReplyReadOption = e.parameter.vmAutoReplyReadOption;
  properties.vmAutoReplyBody = e.parameter.vmAutoReplyBody;
  properties.vmAutoReplyPlayOption = e.parameter.vmAutoReplyPlayOption;
  properties.vmAutoReplyFile = e.parameter.vmAutoReplyFile;
  properties.vmAutoForwardOption = e.parameter.vmAutoForwardOption;
  properties.vmAutoForwardNumber = e.parameter.vmAutoForwardNumber;
  
  properties.vmEnabled = e.parameter.vmEnabled;
  properties.vmTrigger = e.parameter.vmTrigger;
  properties.vmNumSelected = e.parameter.vmNumSelected;
 
  
  var vmPropertyObject = new Object();
  for (var i=0; i<properties.vmNumSelected; i++) {
    vmPropertyObject['vmName-'+i] = e.parameter['vmName-'+i];
    if ((headers.indexOf(vmPropertyObject['vmName-'+i] + " VM Status") == -1)&&(properties.vmEnabled=="true")) {
      sheet.insertColumnAfter(sheet.getLastColumn());
      sheet.getRange(1, sheet.getLastColumn()+1).setValue(vmPropertyObject['vmName-'+i] + " VM Status").setBackgroundColor('orange').setComment('Don\'t change the name of this column. It is used to log the status of an outbound Voice Message, whose name must be a match.');
      sheet.setColumnWidth(sheet.getLastColumn(), "400");
    }
    vmPropertyObject['vmPhone-'+i] = e.parameter['vmPhone-'+i];
    vmPropertyObject['vmBody-'+i] = e.parameter['vmBody-'+i];
    vmPropertyObject['vmSoundFile-'+i] = e.parameter['vmSoundFile-'+i];
    vmPropertyObject['vmRecordOption-'+i] = e.parameter['vmRecordOption-'+i];
    vmPropertyObject['vmLang-'+i] = e.parameter['vmLang-'+i];
    vmPropertyObject['vmCol-'+i] = e.parameter['vmCol-'+i];
    vmPropertyObject['vmVal-'+i] = e.parameter['vmVal-'+i];
  }
  var vmPropertyString = Utilities.jsonStringify(vmPropertyObject);
  properties.vmPropertyString = vmPropertyString;
  ScriptProperties.setProperties(properties);
  app.close();
  return app;
}

function generatearray(range) {
  var output = '[';
  for (var i=0; i<range.length; i++) {
    output += "'" + range[i][0] + "',";
  }
  output += ']';
  return output;
}


function doGet(e) {
  var app = UiApp.createApplication();
  var properties = ScriptProperties.getProperties();
  var accountSID = properties.accountSID;
  var authToken = properties.authToken;
  var twilioNumber = properties.twilioNumber;
  var messageId = e.parameter.MessageId;
  var message = CacheService.getPublicCache().get(messageId);
  var language = e.parameter.Language;
  if (!language) {
    language = properties.defaultVoiceLang;
  }
  var ssId = properties.ssId;
  var ss = SpreadsheetApp.openById(ssId);
  if (e.parameter.Verify == "TRUE") {
    var output = ContentService.createTextOutput();
    var content ='<verified>TRUE</verified>';
     output.setContent(content);
     output.setMimeType(ContentService.MimeType.XML);
     return output;
  }
  if (e.parameter.Body == "Woot woot") {
    var args = new Object();
    args.runTest = "TRUE";
    formMule_sendAText(e.parameter.From, "Congratulations! Your script is able to recieve texts via Twilio. Woot woot to you!", accountSID, authToken, '', args, language);
    formMule_makeRoboCall(e.parameter.From, "Congratulations! Your script is able to create voice messages via Twilio. Hoot Hoot to you!", accountSID, authToken, '', args, language);
    onOpen();
    formMule_smsAndVoiceSettings();
    return;
  }
  
  if (e.parameter.Voice == "TRUE") {
    var output = ContentService.createTextOutput();
    var responseArgs = '';
    if ((message)&&(message!='')) {
      responseArgs += '<Say voice="woman" language="'+language+'">'+message+'</Say>';
    }
    if (e.parameter.PlayFile) {
      responseArgs += '<Play>' + e.parameter.PlayFile + '</Play>';
    }
    if (e.parameter.RequestResponse == "TRUE") {
      responseArgs += '<Record />';
    }
    var content ='<Response>'+responseArgs+'</Response>';
    output.setContent(content);
    output.setMimeType(ContentService.MimeType.XML);
    var flagSheetName = e.parameter.flagSheetName;
    var flagHeader = e.parameter.flagHeader;
    var flagRow = e.parameter.flagRow;
    if ((flagSheetName)&&(flagHeader)&&(flagRow)) {
      var sheet = ss.getSheetByName(flagSheetName);
      var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
      var statusCol = headers.indexOf(unescape(flagHeader))+1
      var range = sheet.getRange(flagRow, statusCol);
      var sid = e.parameter.CallSid;
      var statusObj = getCallStatus(sid);
      updateCallStatus(range, statusObj);
    }
    return output;
  }
  
  if ((e.parameter.VoiceCallBack == "TRUE")&&(e.parameter.OutboundMessageRecording == "TRUE")) {
    var sheet = ss.getSheetByName("MyVoiceRecordings");
    var lastRow = sheet.getLastRow();
    var callSid = e.parameter.CallSid;
    var recordingUrl = getRecordingUrl(callSid);
    var range = sheet.getRange(lastRow+1, 1, 1, 2).setValues([[e.parameter.RecordingName,recordingUrl]]);
  }
  
   if (e.parameter.VoiceCallBack == "TRUE") {
    var flagSheetName = e.parameter.flagSheetName;
    var flagHeader = e.parameter.flagHeader;
    var flagRow = e.parameter.flagRow;
    if ((flagSheetName)&&(flagHeader)&&(flagRow)) {
      var sheet = ss.getSheetByName(flagSheetName);
      var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
      var statusCol = headers.indexOf(unescape(flagHeader))+1
      var range = sheet.getRange(flagRow, statusCol);
      var sid = e.parameter.CallSid;
      var statusObj = getCallStatus(sid);
      var recordingUrl = getRecordingUrl(sid);
      updateCallStatus(range, statusObj, recordingUrl);
    }
     return;
  }
  
  if (e.parameter.CallSid) {
    var output = ContentService.createTextOutput();
    var content = '';
    var callString = '';
    if (properties.vmAutoReplyReadOption=="true") {
      callString += '<Say voice="woman">' + properties.vmAutoReplyBody + '</Say>';
    }
    if (properties.vmAutoReplyPlayOption=="true") {
      callString += '<Play>' + properties.vmAutoReplyFile + '</Play>';
    }
    if (properties.vmAutoForwardOption=="true") {
      callString += '<Dial>' + properties.vmAutoForwardNumber + '</Dial>'
    }
    var content ='<Response>' + callString + '</Response>';
    output.setContent(content);
    output.setMimeType(ContentService.MimeType.XML);
    return output;
  }
  
   if (e.parameter.SmsMessageSid) {
     var output = ContentService.createTextOutput();
     var content = '';
     if (properties.smsAutoReplyOption=="true") {
       content = '<Response><Sms>'+ properties.smsAutoReplyBody + '</Sms></Response>';
     }
     if (properties.smsAutoForwardOption=="true") {
       var email = properties.smsAutoForwardEmail;
       var subj = 'SMS Received from ' + e.parameter.From;
       var body = e.parameter.Body
       MailApp.sendEmail(email, subj, body);
     }
    output.setContent(content);
    output.setMimeType(ContentService.MimeType.XML);
    return output; 
  }
  return app;
}

function updateCallStatus(range, obj, recordingUrl) {
  var app = UiApp.getActiveApplication();
  var ssId = ScriptProperties.getProperty("ssId");
  var ss = SpreadsheetApp.openById(ssId);
  var messages = range.getValue();
  var status = '';
  var sid = obj.Sid.getText();
  status = obj.Status.getText();
  messages = messages.split(";\n");
  var newValues = [];
  var values = [];
  for (var i=0; i<messages.length; i++) {
    values = [];
    values = messages[i].split(", ");
    if (values[1]==("SID="+sid)) {
      values[5] = "Status=" + status;
      if (status=="completed") {
        values[6] = "Duration=" + obj.Duration.getText() + "sec";
        values[7] = "AnsweredBy=" + obj.AnsweredBy.getText();
      }
      if (recordingUrl) {
        values[8] = "RecordingUrl=" + recordingUrl;
      }
      newValues[i] = values.join(", ");
    } else {
      newValues[i] = messages[i];
    }
  }
  newValues = newValues.join(";\n");
  range.setValue(newValues);
  return app;
}


function formMule_splitTexts(longMessage, maxSplits) {
  var charLimit = 160;
  var indicator;
  var breakPoint;
  var messages = [];
  var n;
  var m;
  var numTexts = longMessage.length / charLimit;
  numTexts = Math.ceil(numTexts);
  if (numTexts>maxSplits) {
    numTexts = maxSplits;
  }
  for (n = 0, m = 0; n < numTexts && n < maxSplits; n++) {
    m = n * charLimit;
    // set the indicator so we can now how long it is
    
    if (numTexts>1) {
      indicator = '(' + (n + 1) + '/' + numTexts + ')';
    } else {
      indicator = '';
    }
    // set the breakpoint, taking indicator length into consideration
    breakPoint = m + charLimit - indicator.length;
    // insert the indicator into the correct spot
    longMessage = longMessage.substring(0, breakPoint) + indicator + longMessage.substring(breakPoint);
    // add a message (will be charLimit long and include the indicator)
    messages.push(longMessage.substring(m, m + charLimit));
  }
  return messages;
}


//flagObject contains three parameters: sheetName, header, and row to specify the cell where callback information will be populated
function formMule_sendAText(phoneNumber, message, accountSID, authToken, flagObject, args, lang, maxTexts){
  var ssId = ScriptProperties.getProperty("ssId");
  var ss = SpreadsheetApp.openById(ssId);
  var timeZone = ss.getSpreadsheetTimeZone();
  var triggerFlag = false;
  var flagString = '';
  if (flagObject) {
    var flagSheetName = flagObject.sheetName;
    var flagHeader = flagObject.header;
    var flagRow = flagObject.row;
    flagString = "&flagSheetName="+flagSheetName+"&flagHeader="+flagHeader+"&flagRow="+flagRow;
    var sheet = ss.getSheetByName(flagSheetName);
    var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
    var statusCol = headers.indexOf(unescape(flagHeader))+1
    var range = sheet.getRange(flagRow, statusCol);
  } else {
    flagString = '';
  }
  accountSID = accountSID||ScriptProperties.getProperty("accountSID");
  authToken = authToken||ScriptProperties.getProperty("authToken");
  var twilioNumber = ScriptProperties.getProperty("twilioNumber");
  if (accountSID == null || accountSID == undefined){
    ScriptProperties.setProperty("accountSID", Browser.inputBox("Please Enter your Twilio account SID"))
    accountSid = ScriptProperties.getProperty("accountSID");
  }
  if (authToken == null || authToken == undefined){
    ScriptProperties.setProperty("authToken", Browser.inputBox("Please Enter your Twilio authorization token"))
    authToken = ScriptProperties.getProperty("authToken");
  }
  if (lang) {
    message = LanguageApp.translate(message, "", lang);
  }
  if (!maxTexts) {
    maxTexts = 3;
  }
  message = formMule_splitTexts(message, maxTexts);
  var argGetString = '';
  if (!args) {
    var args = new Object();
  }
  var url = ScriptApp.getService().getUrl();
  phoneNumber = phoneNumber.replace(/\s+/g, ' ');
  var phoneNumberArray = phoneNumber.split(",");
  var status = '';
  for (var j=0; j<phoneNumberArray.length; j++) {
    for (var i=0; i<message.length; i++) {
      args.messageNum = encodeURIComponent((i+1) + " of " + (message.length));
      if (args) {
        for (var key in args) {
          argGetString += "&" + key + "=" + args[key];
        }
      }
      if (message[i]!='') {
      var formatPOST = {
       "From" : twilioNumber,
       "To" : phoneNumberArray[j],
       "Body" : message[i],
       };
      var callObj = {
      method: "POST",
      payload: formatPOST,
      headers: {Authorization: 'Basic ' + Utilities.base64Encode(accountSID + ":" + authToken)},
      contentType: "application/x-www-form-urlencoded", 
      };
      try {
        var text = UrlFetchApp.fetch(https + webName + accountSID + "/SMS/Messages",callObj);
        var response = text.getContentText();
        response = Xml.parse(response);
        var thisStatus = response.TwilioResponse.SMSMessage.Status.getText();
        if (thisStatus=="queued") {
          triggerFlag = true;
        }
        var created = response.TwilioResponse.SMSMessage.DateCreated.getText();
        created = new Date(created);
        created = Utilities.formatDate(created, timeZone, "MM/dd/yy HH:mm:ss");
        var sid = response.TwilioResponse.SMSMessage.Sid.getText();
        status += "SMSMessage" + (i+1) + "/" + message.length + ", SID=" + sid + ", To=" + phoneNumberArray[j] + ", Created=" + created + ", \"" + message[i] + "\", Status=" + thisStatus;
        if ((i+1)<message.length) {
          status += ";\n";
        }
      }
       catch(e) {
        Logger.log(e);
      }
    }
  }
  var currValue = '';
  if ((j+1)<phoneNumberArray.length) {
    status += "|\n";
  }
}
  if (range) {
    range.setValue(currValue + status);
    if (triggerFlag == true) {
      setUpdateTrigger();
    }
  }
  return;
}


function setUpdateTrigger() {
  var triggers = ScriptApp.getScriptTriggers();
  for (var i=0; i< triggers.length; i++) {
    if (triggers[i].getHandlerFunction()=='triggerUpdateStatus') {
      return;
    }
  }
  var now = new Date();
  var newDateObj = new Date(now.getTime() + (0.5)*60000);
  var oneTimeRerunTrigger = ScriptApp.newTrigger('triggerUpdateStatus').timeBased().at(newDateObj).create();
}

function triggerUpdateStatus() {
  var properties = ScriptProperties.getProperties();
  var sheetName = properties.sheetName;
  var smsPropertyString = properties.smsPropertyString;
  if ((smsPropertyString)&&(smsPropertyString!='')) {
    var smsProperties = Utilities.jsonParse(smsPropertyString)
        for (var i=0; i<properties.smsNumSelected; i++) {
          updateStatus(sheetName, smsProperties['smsName-'+i] + " SMS Status");
        }
  }
  var triggers = ScriptApp.getScriptTriggers();
  for (var i=0; i< triggers.length; i++) {
    if (triggers[i].getHandlerFunction()=='triggerUpdateStatus') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}


function updateStatus(sheetName, colName) {
  var triggerFlag = false;
  var ssId = ScriptProperties.getProperty('ssId')
  var ss = SpreadsheetApp.openById(ssId);
  var timeZone = ss.getSpreadsheetTimeZone();
  var sheet = ss.getSheetByName(sheetName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var col = headers.indexOf(colName)+1;
  var range = sheet.getRange(2, col, sheet.getLastRow()-1, 1);
  var values = range.getValues();
  for (var i=0; i<values.length; i++) {
    var value = values[i][0];
    if (value!='') {
      var messages = value.split("|");
      for (var k=0; k<messages.length; k++) {
        var subvalues = messages[k].split(";\n");
        for (var j=0; j<subvalues.length; j++) {
          if (subvalues[j]!='') {
            subvalues[j] = subvalues[j].match(/(".*?"|[^",\s]+)(?=\s*,|\s*$)/g);
            var subsubvalues = subvalues[j] || [];
            var sid = subsubvalues[1].split("=")[1];
            if (subsubvalues[5]=="Status=queued") {
              var statusObj = getSMSStatus(sid);
              var thisStatus = statusObj.status;
              if (thisStatus=="queued") {
                triggerFlag = true;
              }
              subsubvalues[5] = "Status=" + thisStatus;
              if (thisStatus=="sent") {
                var sentTime = statusObj.sent;
                if (sentTime) {
                  sentTime = new Date(sentTime);
                  sentTime = Utilities.formatDate(sentTime, timeZone, "MM/dd/yy HH:mm:ss");
                }
                subsubvalues[6] = "TimeSent=" + sentTime 
              }
            }
          }
          subvalues[j] = subsubvalues.join(", ");
          }
          messages[k] = subvalues.join(";\n");
        }
       values[i][0] = messages.join("|\n"); 
       range.setValues(values)
      }
    }
    if (triggerFlag==true) {
      setUpdateTrigger();
    }
}

function getCallStatus(sid) {
  var accountSID = ScriptProperties.getProperty("accountSID");
  var authToken = ScriptProperties.getProperty("authToken");
  try {
    var fetchUrl = "https://api.twilio.com/2010-04-01/Accounts/"+accountSID+"/Calls/"+sid+".xml";
    var callObj = {
      method: "GET",
      headers: {Authorization: 'Basic ' + Utilities.base64Encode(accountSID + ":" + authToken)}    
    }
        debugger;
    var text = UrlFetchApp.fetch(fetchUrl, callObj);
    var response = text.getContentText();
    response = Xml.parse(response);
    var statusObj = response.TwilioResponse.Call;
    return statusObj;
    }
     catch(e) {
    Logger.log(e);
    }
}

function getRecordingUrl(callSid) {
    var accountSID = ScriptProperties.getProperty("accountSID");
    var authToken = ScriptProperties.getProperty("authToken");
  try {
    var fetchUrl = this.https + this.webName + accountSID + '/Calls/' + callSid + '/Recordings.xml';
    var callObj = {
      method: "GET",
      headers: {Authorization: 'Basic ' + Utilities.base64Encode(accountSID + ":" + authToken)}    
    }
    var text = UrlFetchApp.fetch(fetchUrl, callObj);
    var response = text.getContentText();
    response = Xml.parse(response);
    var recordingSid = response.TwilioResponse.Recordings.Recording[0].getElement().getText();
    var recordingUrl = this.https + this.webName + accountSID + '/Recordings/' + recordingSid;
    return recordingUrl;
  } catch(err) {
     Logger.log(err);
  }
}



function getSMSStatus(sid) {
  var accountSID = ScriptProperties.getProperty("accountSID");
  var authToken = ScriptProperties.getProperty("authToken");
  try {
    var fetchUrl = "https://api.twilio.com/2010-04-01/Accounts/"+accountSID+"/SMS/Messages/"+sid+".xml";
    var callObj = {
      method: "GET",
      headers: {Authorization: 'Basic ' + Utilities.base64Encode(accountSID + ":" + authToken)}    
    }
    var text = UrlFetchApp.fetch(fetchUrl, callObj);
    var response = text.getContentText();
    response = Xml.parse(response);
    var thisStatus = response.TwilioResponse.SMSMessage.Status.getText();
    var thisSent = response.TwilioResponse.SMSMessage.DateSent.getText();
    var statusObj = new Object();
    statusObj.status = thisStatus;
    statusObj.sent = thisSent;
    return statusObj;
    }
     catch(e) {
    Logger.log(e);
    }
}

//flagObject contains three parameters: sheetName, header, and row to specify the cell where callback information will be populated
function formMule_makeRoboCall(phoneNumber, message, accountSID, authToken, flagObject, args, lang) {
  var ssId = ScriptProperties.getProperty("ssId");
  var ss = SpreadsheetApp.openById(ssId);
  var timeZone = ss.getSpreadsheetTimeZone();
  phoneNumber = phoneNumber.replace(/\s+/g, ' ');
  var phoneNumberArray = phoneNumber.split(",");
  var flagString = '';
  if (flagObject) {
    var flagSheetName = flagObject.sheetName;
    var flagHeader = flagObject.header;
    var flagRow = flagObject.row;
    flagString = "&flagSheetName="+flagSheetName+"&flagHeader="+flagHeader+"&flagRow="+flagRow;
  } else {
    flagString = '';
  }
  var triggerFlag = false;
  if (!(lang)) {
    lang = ScriptProperties.getProperty('defaultVoiceLang');
  }
  message = LanguageApp.translate(message, "", lang);
  var validLangs = ['en', 'en-gb', 'es', 'fr', 'de'];
  if (validLangs.indexOf(lang)==-1) {
    lang = 'en';
  }
  var argGetString = '';
  var record = false;
  if ((args)&&(args!='')) {
    for (var key in args) {
      argGetString += "&" + key + "=" + args[key];
    }
  if (args.RequestResponse == "TRUE") {
    record = true;
  }
  }
  var messageId = new Date();
  messageId = messageId.getTime();
  CacheService.getPublicCache().put(messageId, message, 3600);
  accountSID = accountSID||ScriptProperties.getProperty("accountSID");
  authToken = authToken||ScriptProperties.getProperty("authToken");
  var twilioNumber = ScriptProperties.getProperty("twilioNumber");
  var webAppUrl = ScriptApp.getService().getUrl() + "?Voice=TRUE&MessageId="+messageId+"&Language="+lang + flagString + argGetString;
  var callBackUrl = ScriptApp.getService().getUrl() + "?VoiceCallBack=TRUE" + flagString + argGetString;
  var status = '';
  for (var i=0; i<phoneNumberArray.length; i++) {
    var formatPOST = {
      "From" : twilioNumber,
      "To" : phoneNumberArray[i],
      "Url" : webAppUrl,
      "Method" : "GET",
      "StatusCallback" : callBackUrl,
      "StatusCallbackMethod" : "GET",
      "Record" : record,
      "IfMachine" : "Continue"
     };
    var callObj = {
      //work around to get past the inability to use ":" in the urlfetch app
      method: "POST",
      payload: formatPOST,
      headers: {Authorization: 'Basic ' + Utilities.base64Encode(accountSID + ":" + authToken)},    
      contentType: "application/x-www-form-urlencoded",  
      };    
   try{
     var response = UrlFetchApp.fetch(https + webName + accountSID + "/Calls", callObj);
     response = response.getContentText();
     response = Xml.parse(response);
     var thisStatus = response.TwilioResponse.Call.Status.getText();
     if ((thisStatus=="ringing")||(thisStatus=="in-progress")) {
       triggerFlag = true;
     }
     var created = response.TwilioResponse.Call.DateCreated.getText();
     created = new Date(created);
     created = Utilities.formatDate(created, timeZone, "MM/dd/yy HH:mm:ss");
     var sid = response.TwilioResponse.Call.Sid.getText();
     message = message.replace(/,/g, '');
     status += "VoiceMessage, SID=" + sid + ", To=" + phoneNumberArray[i] + ", StartTime=" + created + ", Body=" + message + ", Status=" + thisStatus;
     if ((i+1)<phoneNumberArray.length) {
       status += ";\n";
     }
   }
   catch(err){
     Logger.log(err);
     status += err;
   }
 }
 if (flagSheetName) {
   var sheet = ss.getSheetByName(flagSheetName);
   var headers = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0];
   var statusCol = headers.indexOf(unescape(flagHeader))+1
   var range = sheet.getRange(flagRow, statusCol);
   range.setValue(status);
  }
 return;
}


//gets the associtated phone numbers with the account 
function formMule_getPhoneNumbers(accountSid,authToken) {
  accountSid = accountSid||ScriptProperties.getProperty("twilioID");
  authToken =  authToken||ScriptProperties.getProperty("securityToken");
  var phoneNumbers = [];
//create checks userproreties using 
  if (accountSid == null || accountSid == undefined){
    Browser.msgBox("You are missing your Account SID");
    formMule_howToSetUpTwilio();
  }
  if (authToken == null || authToken == undefined){
    Browser.msgBox("You are missing your Auth Token");
    formMule_howToSetUpTwilio();
  }
      
  var callObj = {
    //work around to get past the inability to use ":" in the urlfetch app
    headers: {Authorization: 'Basic ' + Utilities.base64Encode(accountSid + ":" + authToken)},    
    contentType: "application/xml; charset=utf-8",  
    };

  try{ 
  //change on country will add easy change when fully functional
  //returns an array of JSON objects 1 for each number
    var availableNumbers = UrlFetchApp.fetch(https + webName + accountSid + "/IncomingPhoneNumbers.json", callObj).getContentText();
    availableNumbers = JSON.parse(availableNumbers);
    availableNumbers =  availableNumbers.incoming_phone_numbers;
    if (availableNumbers) {
      for (i=0 ; i < availableNumbers.length ; i++){
        phoneNumbers.push(availableNumbers[i].friendly_name);
      }
    }
    return phoneNumbers;
    }
catch(e){
   Logger.log(e);
   return phoneNumbers;
 }
}

//gets the associtated phone numbers with the account 
function formMule_getOutgoingNumbers(accountSid,authToken) {
  accountSid = accountSid||UserProperties.getProperty("twilioID");
  authToken =  authToken||UserProperties.getProperty("securityToken");
//create checks userproreties using 
  if (accountSid == null || accountSid == undefined){
    UserProperties.setProperty("twilioID", Browser.inputBox("Please Enter your Twilio ID"))
    accountSid = UserProperties.getProperty("twilioID");
    }
  if (authToken == null || authToken == undefined){
    UserProperties.setProperty("securityToken", Browser.inputBox("Please Enter your Twilio autherization key"))
    authToken = UserProperties.getProperty("securityToken");
    }
      
  var callObj = {
    //work around to get past the inability to use ":" in the urlfetch app
    headers: {Authorization: 'Basic ' + Utilities.base64Encode(accountSid + ":" + authToken)},    
    contentType: "application/xml; charset=utf-8",  
    };

  try{ 
  //change on country will add easy change when fully functional
  //returns an array of JSON objects 1 for each number
    var availableNumbers = UrlFetchApp.fetch(https + webName + accountSid + "/OutgoingCallerIds.json", callObj).getContentText();
    availableNumbers = JSON.parse(availableNumbers);
    availableNumbers =  availableNumbers.outgoing_caller_ids;
    var phoneNumbers = [];
    for (i=0 ; i < availableNumbers.length ; i++){
      phoneNumbers.push(availableNumbers[i].friendly_name);
  }
    debugger;
   return phoneNumbers;
    }
catch(e){
    Logger.log(e);
  }
  
}
function formMule_checkForSourceChanges() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  var activeSheetName = activeSheet.getName();
  var activeRange = activeSheet.getActiveCell();
  var activeRow = activeRange.getRow();
  var emailConditionString = ScriptProperties.getProperty('emailConditions');
  if (emailConditionString) {
    var emailConditionObject = Utilities.jsonParse(emailConditionString);
    if (activeRow == 1) {
      var numSelected = ScriptProperties.getProperty("numSelected");
      for (var i = 0; i<numSelected; i++) {
        var sheetName = emailConditionObject['sht-'+i].trim();
        var sheet = ss.getSheetByName(sheetName);
        formMule_setAvailableTags(sheet);
      }
    }
  }
}


// This code was borrowed and modified from the Flubaroo Script author Dave Abouav
// It anonymously tracks script usage to Google Analytics, allowing our non-profit to report our impact to funders
// For original source see http://www.edcode.org

function formMule_institutionalTrackingUi() {
  openInstitutionalTrackingUi();
}


function formMule_logCalEvent()
{
  var systemName = ScriptProperties.getProperty("systemName")
  log("Auto-Created%20Calendar%20Event", scriptName, scriptTrackingId, systemName)
}

function formMule_logCalUpdate()
{
  var systemName = ScriptProperties.getProperty("systemName")
  log("Auto-Updated%20Calendar%20Event", scriptName, scriptTrackingId, systemName)
}


function formMule_logSMS()
{
  var systemName = ScriptProperties.getProperty("systemName")
  log("SMS%20Message%20Sent", scriptName, scriptTrackingId, systemName)
}


function formMule_logVoiceCall()
{
  var systemName = ScriptProperties.getProperty("systemName")
  log("Voice%20Call%20Made", scriptName, scriptTrackingId, systemName)
}


function formMule_logEmail()
{
  var systemName = ScriptProperties.getProperty("systemName")
  log("Mailed%20Templated%20Email", scriptName, scriptTrackingId, systemName)
}


function logRepeatInstall() {
  var systemName = ScriptProperties.getProperty("systemName")
  log("Repeat%20Install", scriptName, scriptTrackingId, systemName)
}

function logFirstInstall() {
  var systemName = ScriptProperties.getProperty("systemName")
  log("First%20Install", scriptName, scriptTrackingId, systemName)
}


function setSid() { 
  var scriptNameLower = scriptName.toLowerCase();
  var sid = ScriptProperties.getProperty(scriptNameLower + "_sid");
  if (sid == null || sid == "")
  {
    var dt = new Date();
    var ms = dt.getTime();
    var ms_str = ms.toString();
    ScriptProperties.setProperty(scriptNameLower + "_sid", ms_str);
    var uid = UserProperties.getProperty(scriptNameLower + "_uid");
    if (uid) {
      logRepeatInstall();
    } else {
      logFirstInstall();
      UserProperties.setProperty(scriptNameLower + "_uid", ms_str);
    }      
  }
}



function formMule_clearAllFlags() {
  formMule_clearMergeFlags();
  formMule_clearEventUpdateFlags();
  formMule_clearEventCreateFlags()
}


function formMule_clearSpecificTemplateMergeFlags() {
  var templateName = "First Reminder";  //Hand code the name of the template whose status you want deleted
  var sheetName = ScriptProperties.getProperty('sheetName');
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var emailConditionString = ScriptProperties.getProperty('emailConditions');  
  var emailConditionObject = Utilities.jsonParse(emailConditionString);    
  var numSelected = emailConditionObject['max'];
  var copyDownOn = ScriptProperties.getProperty('copyDownOption');
  var k = 2;
  if (copyDownOn=="true") {
    k=3;
  }
  var col = headers.indexOf(templateName + " Status") + 1;
  var lastRow = sheet.getLastRow();
  if ((col!=0)&&(lastRow>k-1)) {
    var range = sheet.getRange(k, col, lastRow-(k-1), 1).clear();
  }
}



function formMule_clearMergeFlags() {
  var sheetName = ScriptProperties.getProperty('sheetName');
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName);
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var emailConditionString = ScriptProperties.getProperty('emailConditions');  
  if (emailConditionString) {
    var emailConditionObject = Utilities.jsonParse(emailConditionString);  
    var numSelected = emailConditionObject['max'];
  }
  var copyDownOn = ScriptProperties.getProperty('copyDownOption');
  var k = 2;
  if (copyDownOn=="true") {
    k=3;
  }
  for (var i=0; i<numSelected; i++) {
    var col = headers.indexOf(emailConditionObject['sht-'+i] + " Status") + 1;
    var lastRow = sheet.getLastRow();
    if ((col!=0)&&(lastRow>k-1)) {
      var range = sheet.getRange(k, col, lastRow-(k-1), 1).clear();
    }
  }
}



function formMule_clearEventUpdateFlags() {
  var sheetName = ScriptProperties.getProperty('sheetName');
  var ss = SpreadsheetApp.getActive();
  if ((sheetName)&&(sheetName!='')) {
    var sheet = ss.getSheetByName(sheetName);
    var headers = formMule_fetchHeaders(sheet);
    var eventUpdateIndex = headers.indexOf("Event Update Status");
    var lastRow = sheet.getLastRow();
    var copyDownOn = ScriptProperties.getProperty('copyDownOption');
    var k = 2;
    if (copyDownOn=="true") {
      k=3;
    }
    if ((eventUpdateIndex!=-1)&&(eventUpdateIndex>k-1)) {
      if (lastRow-(k-1)>0) {
        var range = sheet.getRange(k, eventUpdateIndex+1, lastRow-(k-1), 1).clear();
      }
    }
  }
}


function formMule_clearEventCreateFlags() {
  var sheetName = ScriptProperties.getProperty('sheetName');
  var ss = SpreadsheetApp.getActive();
  if ((sheetName)&&(sheetName!='')) {
    var sheet = ss.getSheetByName(sheetName);
    var headers = formMule_fetchHeaders(sheet);
    var eventUpdateIndex = headers.indexOf("Event Creation Status");
    var lastRow = sheet.getLastRow();
    var copyDownOn = ScriptProperties.getProperty('copyDownOption');
    var k = 2;
    if (copyDownOn=="true") {
      k=3;
    }
    if ((eventUpdateIndex!=-1)&&(eventUpdateIndex>k-1)) {
      if (lastRow-(k-1)>0) {
        var range = sheet.getRange(k, eventUpdateIndex+1, lastRow-(k-1), 1).clear();
      }
    }
  }
}

function rangetotable(input, headers) {
  var output = '<table cellpadding="8" cellspacing="0" style="border: 1px solid grey;">';
  var rows = input.length;
  var cols = input[0].length;
  if (headers) {
    if (headers[0].length != cols) {
      return "input and header arrays must be the same width";
    }
    output += "<tr>";
    for (var i=0; i<headers[0].length; i++) {
      output += "<th style=\"background-color:#DDDDDD; border: 1px solid grey;\">"+headers[0][i]+"</th>";
    }
    output += "</tr>";
  }
  for (var i=0; i<input.length; i++) {
    output += "<tr>";
    for (var j=0; j<input[0].length; j++) {
        output += "<td style=\"border: 1px solid grey;\">"+input[i][j]+"</td>";
      }
    output += "</tr>";
  }
  output += "</table>";
  return output;
}


function rangetoverticaltable(input, headers) {
  var output = '<table cellpadding="8" cellspacing="0" style="border: 1px solid grey;">';
  var rows = input[0].length;
  if (headers) {
    if (headers[0].length != rows) {
      return "input and header arrays must be the same width";
    }
  }
  for (var i=0; i<input[0].length; i++) {
    output += "<tr>";
     if (headers) {
      output += "<th style=\"background-color:#DDDDDD; border: 1px solid grey;\">"+headers[0][i]+"</th>";
    }
    output += "<td style=\"border: 1px solid grey;\">"+input[0][i]+"</td>";
    output += "</tr>";
  }
  output += "</table>";
  return output;
}
function formMule_whatIs() {
  var app = UiApp.createApplication().setHeight(550);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var panel = app.createVerticalPanel();
  var muleGrid = app.createGrid(1, 2);
  var image = app.createImage(this.MULEICONURL);
  image.setHeight("100px");
  var label = app.createLabel("formMule: A flexible email, calendar, SMS, and voice merge utility for use with Google Spreadsheet and/or Form data.");
  label.setStyleAttribute('fontSize', '1.5em').setStyleAttribute('fontWeight', 'bold');
  muleGrid.setWidget(0, 0, image);
  muleGrid.setWidget(0, 1, label);
  var mainGrid = app.createGrid(4, 1);
  var html = "<h3>Features</h3>";
      html += "<ul><li>Set up and generate templated, merged emails from form or spreadsheet data.</li>";
      html += "<li>Optionally set 'send conditions' that trigger up to six different emails based on a column's value in the merged row.  Allows for branching or differentiated outputs based on the value of individual form responses.</li>";
      html += "<li>Can be set to auto-copy-down formula columns (on form submit) that operate to the right of form data.  Great for use with VLOOKUP and IF formulas that reference form data.  For example, look up an email address in another sheet based on a name submitted in the form.  This feature is available under \"Advanced options.\"</li>"; 
      html += "<li>Auto-generate and auto-update calendar events using form or spreadsheet data and conditions.</li>";
      html += "<li>Easily connects to the Twilio SMS and Voice service to enable SMS and Voice merges from Spreadsheet or Form data.</li>";
      html += "<li>Can be triggered on form submit or run manually after merge preview.</li>"; 
      html += "<li>Configuration settings can be exported to allow Spreadsheet-based systems to be packaged for easy replication through File->Copy in Google Drive.</li>";
      html += "<li>Interested in doing Google Document or PDF merges instead? Check out the autoCrat script in the gallery.</li></ul>";
    
  mainGrid.setWidget(0, 0, app.createHTML(html));
  var sponsorLabel = app.createLabel("Brought to you by");
  var sponsorImage = app.createImage("http://www.youpd.org/sites/default/files/acquia_commons_logo36.png");
  var supportLink = app.createAnchor('Get the tutorials!', 'http://cloudlab.newvisions.org/scripts/formmule');
  mainGrid.setWidget(1, 0, sponsorLabel);
  mainGrid.setWidget(2, 0, sponsorImage);
  mainGrid.setWidget(3, 0, supportLink);
  app.add(muleGrid);
  panel.add(mainGrid);
  app.add(panel);
  ss.show(app);
  return app;                                                                    
}
function checkInstitutionalTrackingCode(){
 var institutionalTrackingString =  UserProperties.getProperty('institutionalTrackingString');
 if (!institutionalTrackingString){
   closeCurrentUi_()
   institutionalTrackingUi_()
 }
  return institutionalTrackingString;
}

function openInstitutionalTrackingUi() {
  institutionalTrackingUi_()
}

function closeCurrentUi_(){
  var app = UiApp.getActiveApplication(); 
  if (app){
    app.close
    return app
  } else {
    return
  }
}

function institutionalTrackingUi_() {
  closeCurrentUi_();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var institutionalTrackingString = UserProperties.getProperty('institutionalTrackingString');
  var eduSetting = UserProperties.getProperty('eduSetting');
  if (!(institutionalTrackingString)) {
    UserProperties.setProperty('institutionalTrackingString', 'not participating');
  }
  var app = UiApp.createApplication().setTitle('Hello there! Help us track the usage of this script').setHeight(400);
  if ((!(institutionalTrackingString))||(!(eduSetting))) {
    var helptext = app.createLabel("You are most likely seeing this prompt because this is the first time you are using a Google Apps script created by New Visions for Public Schools, 501(c)3. If you are using scripts as part of a school or grant-funded program like New Visions' CloudLab, you may wish to track usage rates with Google Analytics. Entering tracking information here will save it to your user credentials and enable tracking for any other New Visions scripts that use this feature. No personal info will ever be collected.").setStyleAttribute('marginBottom', '10px');
  } else {
    var helptext = app.createLabel("If you are using scripts as part of a school or grant-funded program like New Visions' CloudLab, you may wish to track usage rates with Google Analytics. Entering or modifying tracking information here will save it to your user credentials and enable tracking for any other scripts produced by New Visions for Public Schools, 501(c)3, that use this feature. No personal info will ever be collected.").setStyleAttribute('marginBottom', '10px');
  }
  var absPanel = app.createAbsolutePanel().setHeight(400);
  var panel = app.createVerticalPanel();
  var gridPanel = app.createVerticalPanel().setId("gridPanel").setVisible(false);
  var grid = app.createGrid(4,2).setId('trackingGrid').setStyleAttribute('background', 'whiteSmoke').setStyleAttribute('marginTop', '10px');
  var checkHandler = app.createClientHandler().forTargets(gridPanel).setVisible(true);
  var checkBox = app.createCheckBox('Participate in institutional usage tracking.  (Only choose this option if you know your institution\'s Google Analytics tracker Id.)').setName('trackerSetting').addValueChangeHandler(checkHandler);  
  var checkBox2 = app.createCheckBox('Let New Visions for Public Schools, 501(c)3 know you\'re an educational user.').setName('eduSetting');  
  if ((institutionalTrackingString == "not participating")||(institutionalTrackingString=='')) {
    checkBox.setValue(false);
  } 
  if (eduSetting=="true") {
    checkBox2.setValue(true);
  }
  var institutionNameFields = [];
  var trackerIdFields = [];
  var institutionNameLabel = app.createLabel('Institution Name');
  var trackerIdLabel = app.createLabel('Google Analytics Tracker Id (UA-########-#)');
  grid.setWidget(0, 0, institutionNameLabel);
  grid.setWidget(0, 1, trackerIdLabel);
  if ((institutionalTrackingString)&&((institutionalTrackingString!='not participating')||(institutionalTrackingString==''))) {
    checkBox.setValue(true);
    gridPanel.setVisible(true);
    var institutionalTrackingObject = Utilities.jsonParse(institutionalTrackingString);
  } else {
    var institutionalTrackingObject = new Object();
  }
  for (var i=1; i<4; i++) {
    institutionNameFields[i] = app.createTextBox().setName('institution-'+i);
    trackerIdFields[i] = app.createTextBox().setName('trackerId-'+i);
    if (institutionalTrackingObject) {
      if (institutionalTrackingObject['institution-'+i]) {
        institutionNameFields[i].setValue(institutionalTrackingObject['institution-'+i]['name']);
        if (institutionalTrackingObject['institution-'+i]['trackerId']) {
          trackerIdFields[i].setValue(institutionalTrackingObject['institution-'+i]['trackerId']);
        }
      }
    }
    grid.setWidget(i, 0, institutionNameFields[i]);
    grid.setWidget(i, 1, trackerIdFields[i]);
  } 
  var help = app.createLabel('Enter up to three institutions, with Google Analytics tracker Id\'s.').setStyleAttribute('marginBottom','5px').setStyleAttribute('marginTop','10px');
  gridPanel.add(help);
  gridPanel.add(grid); 
  panel.add(helptext);
  panel.add(checkBox2);
  panel.add(checkBox);
  panel.add(gridPanel);
  
  var imageId = "0B4FHJs5FbHQMNGFaY2hNZVlDajg";
  var url = 'https://drive.google.com/uc?export=download&id='+imageId;
  var image = app.createImage(url).setId("loadingImg").setHeight(150).setWidth(150).setVisible(false)
  
  var button = app.createButton("Save settings");
  var saveDisable = app.createClientHandler().forTargets(button).setEnabled(false).forTargets(image).setVisible(true);
  var saveHandler = app.createServerHandler('saveInstitutionalTrackingInfo').addCallbackElement(panel);
  button.addClickHandler(saveDisable).addClickHandler(saveHandler);
  panel.add(button);
  absPanel.add(panel)
  absPanel.add(image, 70, 100)
  app.add(absPanel);
  ss.show(app);
  return app;
}


function saveInstitutionalTrackingInfo(e) {
  var app = UiApp.getActiveApplication();
  var eduSetting = e.parameter.eduSetting;
  var oldEduSetting = UserProperties.getProperty('eduSetting');
  if (eduSetting == "true") {
    UserProperties.setProperty('eduSetting', 'true');
  }
  if ((oldEduSetting)&&(eduSetting=="false")) {
    UserProperties.setProperty('eduSetting', 'false');
  }
  var trackerSetting = e.parameter.trackerSetting;
  if (trackerSetting == "false") {
    UserProperties.setProperty('institutionalTrackingString', 'not participating');
    app.close();
    return app;
  } else {
    var institutionalTrackingObject = new Object;
    for (var i=1; i<4; i++) {
      var checkVal = e.parameter['institution-'+i];
      if (checkVal!='') {
        institutionalTrackingObject['institution-'+i] = new Object();
        institutionalTrackingObject['institution-'+i]['name'] = e.parameter['institution-'+i];
        institutionalTrackingObject['institution-'+i]['trackerId'] = e.parameter['trackerId-'+i];
        if (!(e.parameter['trackerId-'+i])) {
          Browser.msgBox("You entered an institution without a Google Analytics Tracker Id");
          autoCrat_institutionalTrackingUi()
        }
      }
    }
    var institutionalTrackingString = Utilities.jsonStringify(institutionalTrackingObject);
    UserProperties.setProperty('institutionalTrackingString', institutionalTrackingString);
    app.close();
    return app;
  }
}

//These functions are only used with scripts that are installed on systems packaged by New Visions for Public Schools


function getSystemName() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();
  var file = DocsList.getFileById(ssId);
  var parents = file.getParents();
  var found = false;
  var maxIterations = 1;
  if (parents.length > 1) {
    maxIterations = 3;
  }
  var rootFolderId = DocsList.getRootFolder().getId();
  for (var i=0; i<maxIterations; i++) {
    var thisParent = parents[i];
    if ((parents.length>0)&&(thisParent)) {
      if (thisParent.getId()!=rootFolderId) {
        var theseSpreadsheets = thisParent.getFilesByType('spreadsheet');
        for (var j=0; j<theseSpreadsheets.length; j++) {
          var thisName = theseSpreadsheets[j].getName();
          if (thisName == "Read Me") {
            found = true;
            var readMeSS = SpreadsheetApp.openById(theseSpreadsheets[j].getId());
            break;
          }
        }
      }
    }
    if (found) {
      break;
    } 
  }
  if ((found)&&(readMeSS)) {
    var sheet = readMeSS.getSheets()[0];
    var timeZone = readMeSS.getSpreadsheetTimeZone();
    var range = sheet.getRange(1,2,10,1);
    var thisSystem = getColumnsData(sheet, range)[0];
    var version = '';
    if (!thisSystem.systemName) {
      return;
    }
    if (thisSystem.version) {
      version = " - V" + thisSystem.version;
    }
    var dateOfLastUpdate = '';
    if (thisSystem.dateOfLastUpdate) {
      dateOfLastUpdate = " (" + Utilities.formatDate(new Date(thisSystem.dateOfLastUpdate), timeZone, 'M/dd/yy') + ")";
    }
    var trackingName = thisSystem.systemName + version + dateOfLastUpdate;
    return trackingName;
  } else {
    return;
  } 
}


// getSheetById looks within a spreadsheet for a sheet with the matching ID
// Arguments:
//    - sheetId: the sheet id of the sheet we're looking for.  Should be an integer but
//         as a safeguard this function will attempt to convert numeric strings to numbers.
//    - spreadsheet: the spreadsheet object that contains the sheets of interest
// Returns a sheet object.  Throws a Browser message box to user if sheet isn't found.
/*
* @param {sheetId} integer or numeric string sheet Id
* @param {spreadsheet} spreadsheet where we want to find the sheet
*/

function getSheetById(spreadsheet, sheetId) {
  try {
    //DO NOT NEED THIS IN NEW SHEETS
    //sheetId = parseInt(sheetId);
    var sheets = spreadsheet.getSheets();
    for (var i=0; i<sheets.length; i++) {
      if (sheets[i].getSheetId() == sheetId) {
        return sheets[i];
      }
    }
    debugger;
    Browser.msgBox("Sheet with ID " + sheetId + " not found. You may have deleted a sheet this script depends on.");
    return;
  } catch(err) {
    throw(err.message);
    return;
  }
}


// getRowsData iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//       This argument is optional and it defaults to all the cells except those in the first row
//       or all the cells below columnHeadersRowIndex (if defined).
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
// Returns an Array of objects.
/*
 * @param {sheet} sheet with data to be pulled from.
 * @param {range} range where the data is in the sheet, headers are above
 * @param {row} 
 */
function getRowsData(sheet, range, columnHeadersRowIndex) {
  var headersIndex = columnHeadersRowIndex || range ? range.getRowIndex() - 1 : 1;
  if (columnHeadersRowIndex) {
    headersIndex = columnHeadersRowIndex;
  }
  var dataRange = range ||
    sheet.getRange(headersIndex + 1, 1, sheet.getMaxRows() - headersIndex, sheet.getMaxColumns());
  var numColumns = dataRange.getLastColumn() - dataRange.getColumn() + 1;
  var headersRange = sheet.getRange(headersIndex, dataRange.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects_(dataRange.getValues(), normalizeHeaders(headers));
}


// getRowsFormulas iterates row by row in the input range and returns an array of objects.
// Each object contains all the data for a given row, indexed by its normalized column name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//       This argument is optional and it defaults to all the cells except those in the first row
//       or all the cells below columnHeadersRowIndex (if defined).
//   - columnHeadersRowIndex: specifies the row number where the column names are stored.
//       This argument is optional and it defaults to the row immediately above range;
// Returns an Array of objects.
/*
 * @param {sheet} sheet with data to be pulled from.
 * @param {range} range where the data is in the sheet, headers are above
 * @param {row} 
 */
function getRowsFormulas(sheet, range, columnHeadersRowIndex) {
  var headersIndex = columnHeadersRowIndex || range ? range.getRowIndex() - 1 : 1;
  var dataRange = range ||
    sheet.getRange(headersIndex + 1, 1, sheet.getMaxRows() - headersIndex, sheet.getMaxColumns());
  var numColumns = dataRange.getLastColumn() - dataRange.getColumn() + 1;
  var headersRange = sheet.getRange(headersIndex, dataRange.getColumn(), 1, numColumns);
  var headers = headersRange.getValues()[0];
  return getObjects_(dataRange.getFormulas(), normalizeHeaders(headers));
}


// getColumnsData iterates column by column in the input range and returns an array of objects.
// Each object contains all the data for a given column, indexed by its normalized row name.
// Arguments:
//   - sheet: the sheet object that contains the data to be processed
//   - range: the exact range of cells where the data is stored
//   - rowHeadersColumnIndex: specifies the column number where the row names are stored.
//       This argument is optional and it defaults to the column immediately left of the range;
// Returns an Array of objects.
function getColumnsData(sheet, range, rowHeadersColumnIndex) {
  rowHeadersColumnIndex = rowHeadersColumnIndex || range.getColumnIndex() - 1;
  var headersTmp = sheet.getRange(range.getRow(), rowHeadersColumnIndex, range.getNumRows(), 1).getValues();
  var headers = normalizeHeaders(arrayTranspose(headersTmp)[0]);
  return getObjects_(arrayTranspose(range.getValues()), headers);
}


// For every row of data in data, generates an object that contains the data. Names of
// object fields are defined in keys.
// Arguments:
//   - data: JavaScript 2d array
//   - keys: Array of Strings that define the property names for the objects to create
function getObjects_(data, keys) {
  var objects = [];
  for (var i = 0; i < data.length; ++i) {
    var object = {};
    var hasData = false;
    for (var j = 0; j < data[i].length; ++j) {
      var cellData = data[i][j];
      if (isCellEmpty_(cellData)) {
        object[keys[j]] = '';
        continue;
      }
      object[keys[j]] = cellData;
      hasData = true;
    }
    if (hasData) {
      objects.push(object);
    }
  }
  return objects;
}

// Returns an Array of normalized Strings.
// Empty Strings are returned for all Strings that could not be successfully normalized.
// Arguments:
//   - headers: Array of Strings to normalize
function normalizeHeaders(headers) {
  var keys = [];
  for (var i = 0; i < headers.length; ++i) {
    keys.push(normalizeHeader(headers[i]));
  }
  return keys;
}

// Normalizes a string, by removing all alphanumeric characters and using mixed case
// to separate words. The output will always start with a lower case letter.
// This function is designed to produce JavaScript object property names.
// Arguments:
//   - header: string to normalize
// Examples:
//   "First Name" -> "firstName"
//   "Market Cap (millions) -> "marketCapMillions
//   "1 number at the beginning is ignored" -> "numberAtTheBeginningIsIgnored"
function normalizeHeader(header) {
  var key = "";
  var upperCase = false;
  for (var i = 0; i < header.length; ++i) {
    var letter = header[i];
    if (letter == " " && key.length > 0) {
      upperCase = true;
      continue;
    }
    if (!isAlnum_(letter)) {
      continue;
    }
    if (key.length == 0 && isDigit_(letter)) {
      continue; // first character must be a letter
    }
    if (upperCase) {
      upperCase = false;
      key += letter.toUpperCase();
    } else {
      key += letter.toLowerCase();
    }
  }
  return key;
}

// Returns true if the cell where cellData was read from is empty.
// Arguments:
//   - cellData: string
function isCellEmpty_(cellData) {
  return typeof(cellData) == "string" && cellData == "";
}

// Returns true if the character char is alphabetical, false otherwise.
function isAlnum_(char) {
  return char >= 'A' && char <= 'Z' ||
    char >= 'a' && char <= 'z' ||
    isDigit_(char);
}

// Returns true if the character char is a digit, false otherwise.
function isDigit_(char) {
  return char >= '0' && char <= '9';
}

// setRowsData fills in one row of data per object defined in the objects Array.
// For every Column, it checks if data objects define a value for it.
// Arguments:
//   - sheet: the Sheet Object where the data will be written
//   - objects: an Array of Objects, each of which contains data for a row
//   - optHeadersRange: a Range of cells where the column headers are defined. This
//     defaults to the entire first row in sheet.
//   - optFirstDataRowIndex: index of the first row where data should be written. This
//     defaults to the row immediately below the headers.
function setRowsData(sheet, objects, optHeadersRange, optFirstDataRowIndex) {
  var headersRange = optHeadersRange || sheet.getRange(1, 1, 1, sheet.getMaxColumns());
  var firstDataRowIndex = optFirstDataRowIndex || headersRange.getRowIndex() + 1;
  var headers = normalizeHeaders(headersRange.getValues()[0]);
  var data = [];
  for (var i = 0; i < objects.length; ++i) {
    var values = []
    for (j = 0; j < headers.length; ++j) {
      var header = headers[j];
      // If the header is non-empty and the object value is 0...
       if ((header.length > 0)&&(objects[i][header] === 0)&&(!(isNaN(parseInt(objects[i][header]))))) {
        values.push(0);
      }
      // If the header is empty or the object value is empty...
      else if ((!(header.length > 0)) || (objects[i][header]=='') || (!objects[i][header])) {
        values.push('');
      }
      else {
        values.push(objects[i][header]);
      }
    }
    data.push(values);
  }

  var destinationRange = sheet.getRange(firstDataRowIndex, headersRange.getColumnIndex(),
                                        objects.length, headers.length);
  destinationRange.setValues(data);
}



// Given a JavaScript 2d Array, this function returns the transposed table.
// Arguments:
//   - data: JavaScript 2d Array
// Returns a JavaScript 2d Array
// Example: arrayTranspose([[1,2,3],[4,5,6]]) returns [[1,4],[2,5],[3,6]].
function arrayTranspose(data) {
  if (data.length == 0 || data[0].length == 0) {
    return null;
  }

  var ret = [];
  for (var i = 0; i < data[0].length; ++i) {
    ret.push([]);
  }

  for (var i = 0; i < data.length; ++i) {
    for (var j = 0; j < data[i].length; ++j) {
      ret[j][i] = data[i][j];
    }
  }

  return ret;
}
/*
* @param {String} - action to log i.e firstInstal
* @param {String} - script name  i.e autoCrat
* @param {String} - system name, optonal
*/
function log(action, scriptName, scriptTrackingId, systemName)
{
  var ga_url = createGATrackingUrl_(action, scriptTrackingId, scriptName);
  if (ga_url)
    {
      var response = UrlFetchApp.fetch(ga_url);
    }
  var institutionalTrackingObject = getInstitutionalTrackerObject_();
  if (institutionalTrackingObject) {
    if (systemName) {
      var encoded_system_name = urlencode_(systemName);
      createSystemTrackingUrls_(institutionalTrackingObject, encoded_system_name, scriptName, action);
      createInstitutionalTrackingUrls_(institutionalTrackingObject, encoded_system_name, action, scriptName); //test isn't needed only the call to createInstitutionalTrackingUrls_
    } else {
      createInstitutionalTrackingUrls_(institutionalTrackingObject, '', action, scriptName); 
    }
  }
}


function createGATrackingUrl_(encoded_page_name, scriptTrackingId, scriptName)
{
  var utmcc = createGACookie_(scriptName);
  var eduSetting = UserProperties.getProperty('eduSetting');
   if (eduSetting=="true") {
    encoded_page_name = "edu/" + encoded_page_name;
  }
  if (utmcc == null)
    {
      return null;
    }
 
  var ga_url1 = "http://www.google-analytics.com/__utm.gif?utmwv=5.2.2&utmhn=www.newvisions-analytics.com&utmcs=-&utmul=en-us&utmje=1&utmdt&utmr=0=";
  var ga_url2 = "&utmac=" + scriptTrackingId + "&utmcc=" + utmcc + "&utmu=DI~"; //Id Changed
  var ga_url_full = ga_url1 + encoded_page_name + "&utmp=" + encoded_page_name + ga_url2;
  
  return ga_url_full;
}

// Some of this code was borrowed and modified from the Flubaroo Script author Dave Abouav
// It anonymously tracks script usage to Google Analytics, allowing our non-profit organization to report the impact of this work to funders
// For original source see http://www.edcode.org
function createInstitutionalTrackingUrls_(institutionTrackingObject, encoded_system_name, encoded_page_name, encoded_script_name) {
  for (var key in institutionTrackingObject) {
    var utmcc = createGACookie_(encoded_script_name);
    if (utmcc == null)
    {
      return null;
    }
    if ((encoded_system_name)&&(encoded_system_name!=='')) {
      var encoded_page_name = encoded_system_name+"/"+encoded_script_name+"/"+encoded_page_name;
    } else {
      var encoded_page_name = encoded_script_name+"/"+encoded_page_name;
    }
    var trackingId = institutionTrackingObject[key].trackerId;
    var ga_url1 = "http://www.google-analytics.com/__utm.gif?utmwv=5.2.2&utmhn=www.autocrat-analytics.com&utmcs=-&utmul=en-us&utmje=1&utmdt&utmr=0=";
    var ga_url2 = "&utmac="+trackingId+"&utmcc=" + utmcc + "&utmu=DI~";
    var ga_url_full = ga_url1 + encoded_page_name + "&utmp=" + encoded_page_name + ga_url2;
    
    if (ga_url_full)
    {
      var response = UrlFetchApp.fetch(ga_url_full);
      return ga_url_full
    }
  }
}

function createSystemTrackingUrls_(institutionTrackingObject, encoded_system_name, scriptName, encoded_execution_name) {
  for (var key in institutionTrackingObject) {
    var utmcc = createGACookie_(scriptName);
    if (utmcc == null)
    {
      return null;
    }
    var trackingId = institutionTrackingObject[key].trackerId;
    var encoded_page_name = encoded_system_name+"/"+encoded_execution_name+"/"+trackingId;
    var ga_url1 = "http://www.google-analytics.com/__utm.gif?utmwv=5.2.2&utmhn=www.cloudlab-systems-analytics.com&utmcs=-&utmul=en-us&utmje=1&utmdt&utmr=0=";
    var ga_url2 = "&utmac=UA-34521561-1&utmcc=" + utmcc + "&utmu=DI~";
    var ga_url_full2 = ga_url1 + encoded_page_name + "&utmp=" + encoded_page_name + ga_url2;  
    if (ga_url_full2)
    {
      var response = UrlFetchApp.fetch(ga_url_full2);
    }
  }
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//-------------------------------------------------- Tracking Utilities ----------------------------------------//
/////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function createGACookie_(scriptName)
{
  var a = "";
  var b = "100000000";
  var c = "200000000";
  var d = "";

  var dt = new Date();
  var ms = dt.getTime();
  var ms_str = ms.toString();
 
  var newVisions_uid = UserProperties.getProperty(scriptName.toLowerCase() + "_uid");
  if ((newVisions_uid == null) || (newVisions_uid == ""))
    {
      // shouldn't happen unless user explicitly removed _uid from properties.
      newVisions_uid = setUid_(scriptName.toLowerCase())
    }
  
  a = newVisions_uid.substring(0,9);
  d = newVisions_uid.substring(9);
  
  utmcc = "__utma%3D451096098." + a + "." + b + "." + c + "." + d 
          + ".1%3B%2B__utmz%3D451096098." + d + ".1.1.utmcsr%3D(direct)%7Cutmccn%3D(direct)%7Cutmcmd%3D(none)%3B";
 
  return utmcc;
}

function setUid_(scriptName)
{ 
  var newVisions_uid = UserProperties.getProperty(scriptName.toLowerCase() + "_uid");
  if (newVisions_uid == null || newVisions_uid == "")
    {
      // user has never installed autoCrat before (in any spreadsheet)
      var dt = new Date();
      var ms = dt.getTime();
      var ms_str = ms.toString();
 
      UserProperties.setProperty(scriptName.toLowerCase() + "_uid", ms_str);
    }
  return ms_str
}


function getInstitutionalTrackerObject_() {
  var institutionalTrackingString = UserProperties.getProperty('institutionalTrackingString');
  if ((institutionalTrackingString)&&(institutionalTrackingString != "not participating")) {
    var institutionTrackingObject = Utilities.jsonParse(institutionalTrackingString);
    return institutionTrackingObject;
  }
  return;
}

function urlencode_(inNum) {
  // Function to convert non URL compatible characters to URL-encoded characters
  var outNum = 0;     // this will hold the answer
  outNum = escape(inNum); //this will URL Encode the value of inNum replacing whitespaces with %20, etc.
  return outNum;  // return the answer to the cell which has the formula
}
