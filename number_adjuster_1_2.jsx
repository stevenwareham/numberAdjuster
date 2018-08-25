// numberAdjuster.jsx  Version 1.1.1

/*
	This InDesign CS3/CS4 JavaScript is intended to perform operations on found numbers. It uses the find/change Grep functions
	to find numbers within either the entire document, or just the selection. You can select to search for just numbers with 
	a specific character style, and/or a prefix character.
	
	This script was written by Steve Wareham at stevewareham.com based on an idea by David Blatner at indesignsecrets.com
	*/

//====================================================\\
 // This function removes non-numeric characters, such as the comma
function stripNonNumeric(str)
 {
	str += '';
    var rgx = /^\d|\.|-$/;
     var out = '';
     for( var i = 0; i < str.length; i++ )
     {
       if( rgx.test( str.charAt(i) ) ){
         if( !( ( str.charAt(i) == '.' && out.indexOf( '.' ) != -1 ) ||
                ( str.charAt(i) == '-' && out.length != 0 ) ) ){
           out += str.charAt(i);
         }
       }
     }
     return out;
   }

//====================================================\\
// This function adds commas to numbers
function addCommas(nStr)
   {
     nStr += '';
     x = nStr.split('.');
     x1 = x[0];
     x2 = x.length > 1 ? '.' + x[1] : '';
     var rgx = /(\d+)(\d{3})/;
     while (rgx.test(x1)) {
      x1 = x1.replace(rgx, '$1' + ',' + '$2'); 
    }
    return x1 + x2;
 }

//====================================================\\
//Function to round numbers
function numRounder(myNumber, myRoundto)
{
	myRoundMod = 1;
	for(rounding = 0; rounding < myRoundto; rounding++)	
		{
			myRoundMod = myRoundMod * 10;	
			}
	myRoundedNum = Math.round(myNumber*myRoundMod)/myRoundMod;
return myRoundedNum;
}

//========================================================\\
//Function to get an array of character style names
function myGetCharacterStyleNames(){
	var myStyleNames = app.documents.item(0).characterStyles.everyItem().name;
	return myStyleNames;
}

//========================================================\\
//Check if a preferences file exists, if it does try to read in values
var thisScript = app.activeScript;
var myTempFile = new File (thisScript.path +"/" + 'num_adjust_prefs.txt');
var myTempFileExists = myTempFile.open( 'r', 'TEXT', 'R*ch');
if (myTempFileExists == true)
	{
		try //Try to read the temp file
			{ 
			myTempFileContents = myTempFile.read();
			myTempFile.close;
			myDialogValues = myTempFileContents.split("@,@");
			myOperator = parseFloat(myDialogValues[0]); 
			myOperation = myDialogValues[1];
			switch(myOperation)  
				{
					case "+":	myOperationIndex = 0		break;
					case "-":	myOperationIndex = 1		break;
					case "*":	myOperationIndex = 2		break;			
					case "/" :	myOperationIndex = 3		break;	
					default: 	myOperationIndex = 0		break;
					}				
			mySearchPrefix = myDialogValues[2];
			myCharStyleIndex = parseInt(myDialogValues[3]);
			myCharStyleCheck = myDialogValues[4];
			if(myCharStyleCheck == "true")
				{myCharStyleCheck = true;}
				else {myCharStyleCheck = false;}
			myScopeDocumentRadio = myDialogValues[5];	
			if(myScopeDocumentRadio == "true")
				{myScopeDocumentRadio = true;}
				else {myScopeDocumentRadio = false;}
			myScopeSelectionRadio = myDialogValues[6];	
			if(myScopeSelectionRadio == "true")
				{myScopeSelectionRadio = true;}
				else {myScopeSelectionRadio = false;}
			myScopeStoryRadio = myDialogValues[12];	
			if(myScopeStoryRadio == "true")
				{myScopeStoryRadio= true;}
				else {myScopeStoryRadio = false;}
			myNFASelectionRadio = myDialogValues[7];
			if(myNFASelectionRadio == "true")
				{myNFASelectionRadio = true;}
				else {myNFASelectionRadio = false;}
			myNFESelectionRadio = myDialogValues[8];
			if(myNFESelectionRadio == "true")
				{myNFESelectionRadio = true;}
				else {myNFESelectionRadio = false;}
			myRoundtoDefault = parseInt(myDialogValues[9]);
			myCustomGrepCheck = myDialogValues[10];
			if(myCustomGrepCheck  == "true")
				{myCustomGrepCheck  = true;}
				else {myCustomGrepCheck = false;}
			myCustomGrepFieldContents = myDialogValues[11];	
		}
		catch(e)  // If there was an error with reading pref. file
			{
				myOperator = 1;
				myOperationIndex = 0;
				mySearchPrefix ="";
				myCharStyleIndex = 0;
				myCharStyleCheck = false;
				myScopeStoryRadio = false;
				myScopeDocumentRadio = false;
				myScopeSelectionRadio = true;
				myNFASelectionRadio = true;
				myNFESelectionRadio = false;
				myRoundtoDefault = 2;
				myCustomGrepCheck = false;
				myCustomGrepFieldContents = "[-]*[0-9][0-9,.]*\\b";
				}
		} //End of if there is a temp file
	else{ //Else if the prefs/temp file could not be found
		myOperator = 1;
		myOperationIndex = 0;
		mySearchPrefix ="";
		myCharStyleIndex = 0;
		myCharStyleCheck = false;
		myScopeStoryRadio = false;
		myScopeDocumentRadio = false;
		myScopeSelectionRadio = true;
		myNFASelectionRadio = true;
		myNFESelectionRadio = false;
		myRoundtoDefault = 2;
		myCustomGrepCheck = false;
		myCustomGrepFieldContents = "[-]*[0-9][0-9,.]*\\b";
		}


//========================================================\\
//Display dialog box to with options to user
var myDialog = app.dialogs.add({name:"Number Adjuster 1.1", canCancel:true});
var myOperations = new Array("Addition", "Subtraction", "Multiplication", "Division");
var myCharStyles = myGetCharacterStyleNames();

try{ //try to create the dialog box
	with(myDialog)
		{
			with(dialogColumns.add())
				{
					with(dialogColumns.add())
						{
							with(borderPanels.add())
								{
									with(dialogColumns.add())
										{
											staticTexts.add({staticLabel:"Perform operation on:       "});
											with(mySearchScopeButtons = radiobuttonGroups.add())
												{
													myRadio1 = radiobuttonControls.add({staticLabel:"Entire document", checkedState:myScopeDocumentRadio});
													myRadioStory = radiobuttonControls.add({staticLabel:"Story of selected frame", checkedState:myScopeStoryRadio}); 	
													myRadio2 = radiobuttonControls.add({staticLabel:"Current selection", checkedState:myScopeSelectionRadio}); 											
													}
													staticTexts.add({staticLabel:" "});
													staticTexts.add({staticLabel:"Numbers are formatted as:"});
													with(myNumberFormattButtons = radiobuttonGroups.add())
														{
															myRadio3 = radiobuttonControls.add({staticLabel:"1,234.56", checkedState:myNFASelectionRadio});
															myRadio4 = radiobuttonControls.add({staticLabel:"1.234,56", checkedState:myNFESelectionRadio}); 											
															}	
													staticTexts.add({minWidth:15});	
											}
									}
							}		
					with(dialogColumns.add())
						{
							with(borderPanels.add())
								{		
									staticTexts.add({staticLabel:"Operation to perform:"});
									var myOperationsDropDown = dropdowns.add ({stringList:myOperations, selectedIndex:myOperationIndex });
									staticTexts.add({staticLabel:"Operator: "});
									var myOperatorField = realEditboxes.add({editValue:myOperator});
									}
							with(borderPanels.add())
								{
									staticTexts.add({staticLabel:"How many places to round to?"});
									var myRoundto = realEditboxes.add({editValue:myRoundtoDefault});	
									staticTexts.add({minWidth:125});
									}	
							with(borderPanels.add())
								{
									staticTexts.add({staticLabel:"Only find numbers with prefix character($,#,etc.): "});
									var myPreCharField = textEditboxes.add({editContents:mySearchPrefix});	
									staticTexts.add({minWidth:15});
									}
							with(borderPanels.add())
								{
									var myCharStyleCheckBox = checkboxControls.add();
									myCharStyleCheckBox.checkedState = myCharStyleCheck;
									staticTexts.add({staticLabel:"Only adjust numbers with character style:"});
									var myCharStyleDropDown = dropdowns.add({stringList:myCharStyles, selectedIndex:myCharStyleIndex});	
									staticTexts.add({minWidth:25});
									}		
							with(borderPanels.add())
								{	
								var myCustomGrepCheckBox = checkboxControls.add();
									myCustomGrepCheckBox.checkedState = myCustomGrepCheck;
									staticTexts.add({staticLabel:"Search on Custom RegEx:"});
									var myGrepField = textEditboxes.add({editContents:myCustomGrepFieldContents, minWidth:155});	
									staticTexts.add({minWidth:45});
									
									}						
							}	
					}
			}  // End of the dialog box
	}//End of try
catch(e){alert(e + " Try deleting the preferences file");}

//========================================================\\
//========================================================\\

if(myDialog.show() == true)
	{
		var myRunScript = true; //Key to check if the search and replace funciton should run
		
		//========================================================\\
		//Check length of prefix character, it cannot be longer than 1 character
		var myPrefixcheck = true;
		if(myPreCharField.editContents.length > 1)
			{
				alert("The search prefix cannot be longer than one character.");
				myRunScript = false;
				}		
		
		//========================================================\\
		 //Set the variables, based on what was entered in the dialog box
		var myDocument = app.activeDocument;
		var myNumberFormatIndex = myNumberFormattButtons.selectedButton; 		
		var mySearchScopeIndex =mySearchScopeButtons.selectedButton;   
		switch(mySearchScopeIndex)  //Set the scope used to search on 
			{
				case 0:	myScope = "doc"		break;
				case 1:	myScope = "story"		break;	
				case 2:	myScope = "selection"	break;
				default: 	myScope = "selection"	break;
				}
					
		var myOperator = myOperatorField.editContents; //Change the myOperfield to a realedit box
		myOperator = parseFloat(myOperator); //this wouldnt be needed
					
		var myOperationIndex = myOperationsDropDown.selectedIndex; 
		switch(myOperationIndex)  // Set the operation to perform
			{
				case 0:	myOperation = "+"	break;
				case 1:	myOperation = "-"		break;
				case 2:	myOperation = "*"	break;			
				case 3:	myOperation = "/"		break;	
				default: 	myOperation = "+"	break;
				}		
				
		var myRoundto = parseFloat(myRoundto.editContents);  //change this to a realedit box or combobox
		
		//Get the Prefix character if one is being used for search. Escape RegEx reserved chars. if they are being used.
		var mySearchPrefix = myPreCharField.editContents;  
				
		if(mySearchPrefix == "$" || mySearchPrefix =="[" || mySearchPrefix =="^" || mySearchPrefix =="." || mySearchPrefix == "|" || mySearchPrefix =="?")
			{mySearchPrefix = "\\" + mySearchPrefix;} //Escape special characters search prefixes			
				
		if(mySearchPrefix == "\\" || mySearchPrefix =="+" || mySearchPrefix == "*" ||mySearchPrefix =="(" || mySearchPrefix == ")")
			{mySearchPrefix = "\\" + mySearchPrefix;} //Escape even more special characters search prefixes
			
		if(mySearchPrefix != "") // Put the Prefix in [ ] if it is used, create replacePrefix var
			{myReplacePrefix = mySearchPrefix; mySearchPrefix = "[" + mySearchPrefix + "]";}
			else {myReplacePrefix = mySearchPrefix;}
						
		var myCharStyleIndex =myCharStyleDropDown.selectedIndex; // Get character style to search on
		var myCharStyle = app.documents.item(0).characterStyles[myCharStyleIndex]; 
				
		//Set the Grep to search on
		if (myCustomGrepCheckBox.checkedState == true )
			{myGrep = myGrepField.editContents;}
			else {myGrep = "[-]*[0-9][0-9,.]*\\b";} //else use this default Grep to search on
					
		//========================================================\\
		//Set variables based on which scope to use
		mySearchScope = new Array();	
		
		if (myScope == "doc") 
			{	
				mySearchScope[0] = myDocument;
				mySearchAmount = 1;
				}
			else if (myScope == "story")
				{
					if (app.selection.length != 1)
						{
							myRunScript = false; 
							alert(myErrorMessage = "Select a single text frame and try again");
							}
						else
							{
								mySearchScope[0] = app.selection[0].parentStory;
								mySearchAmount = 1;
								}
					}
			else //Else just search on the selection
				{ 
					mySearchScope = app.selection;
					mySearchAmount = mySearchScope.length;
					}
			
			//========================================================\\
			//Run the search and replace
			if (myRunScript == true)
				{
					//For the amount of top level items in the scope (i.e. selections)	
					for (i = 0; i < mySearchAmount; i++)
						{
							//Clear the find/change preferences.
							app.findGrepPreferences = NothingEnum.nothing;
							app.changeGrepPreferences = NothingEnum.nothing;
									
							try //Try to do the search + perform operation + replace
								{						
								//Search on character style if that option was checked
								if (myCharStyleCheckBox.checkedState == true )
									{app.findGrepPreferences.appliedCharacterStyle = myCharStyle; }
											
								//Find matching numbers
								var mySearchGrep = app.findGrepPreferences.findWhat = mySearchPrefix + myGrep; //Find all numbers
								var myFoundArray = mySearchScope[i].findGrep(); //Array of matches to the search grep																	
								
								//Change matches to the result of the operation								
								for(k = myFoundArray.length - 1; k >= 0; k--)
									{										
										var myFoundNumber = myFoundArray[k].contents;
										if(mySearchPrefix != "")
											{myFoundNumber = myFoundArray[k].contents.slice(1);} //Slice of the search prefix if one was used		
													
										if(myNumberFormatIndex == 1) //If the number format is 1.234,56
											{
												//Switch commas with decimals so we can still do some math on found numbers
												myFoundNumber = myFoundNumber.replace(/,/gi, "comma"); 
												myFoundNumber= myFoundNumber.replace(/\./gi, ","); 
												myFoundNumber = myFoundNumber.replace(/comma/gi, "."); 	
												}
												
										//Strip Commas from the number
										myStrippedFoundNumber = stripNonNumeric(myFoundNumber);
																
										//Perform the selected operation on the found number
										if (myOperation == "+"){var myResult =  (parseFloat(myStrippedFoundNumber) + myOperator) + "";}
											else if (myOperation == "-"){var myResult =  (parseFloat(myStrippedFoundNumber) - myOperator) + "";}
												else if (myOperation == "*"){var myResult =  (parseFloat(myStrippedFoundNumber) * myOperator) + "";}
													else if (myOperation == "/"){var myResult =  (parseFloat(myStrippedFoundNumber) / myOperator) + "";}						
																											
										//Round the result 
										myResult = numRounder(myResult, myRoundto);						
																															
										//Add Comma to the result of the operation		
										myResult = addCommas(myResult);

										if(myNumberFormatIndex == 1)  //If the number format is 1.234,56
											{
												//Switch commas and decimals back to how they were found
												myResult = myResult.replace(/,/gi, "comma"); 
												myResult = myResult.replace(/\./gi, ","); 
												myResult = myResult.replace(/comma/gi, "."); 
												
												myFoundNumber = myFoundNumber.replace(/,/gi, "comma"); 
												myFoundNumber= myFoundNumber.replace(/\./gi, ","); 
												myFoundNumber = myFoundNumber.replace(/comma/gi, "."); 	
												}									
										
										//If the search prefix had a "/" added to it to escape reserved regex chars. remove it from the replace prefix
										if(myReplacePrefix.length > 1)
											{myReplacePrefix = myReplacePrefix.slice(1)}

										//Change the found text to the result of the operation												
										app.changeGrepPreferences.changeTo = (myReplacePrefix + myResult);
										app.findGrepPreferences.findWhat = (mySearchPrefix + myFoundNumber);
										myFoundArray[k].changeGrep();	
										} //End of looping through all found matches, performing operation, and replacing with the result
								} //End of try
						catch(e){alert(e)}
						//Reset Grep Preferences
						app.findGrepPreferences = NothingEnum.nothing;
						app.changeGrepPreferences = NothingEnum.nothing;	
						} //End of looping through all top level objects
				} //End of if myRunScript is true
		
		//========================================================\\
		//Write new values from dialog box into temp file	
		try
			{
				var myTempFileContents = myOperator + "@,@" + myOperation + "@,@" + myPreCharField.editContents + "@,@" + myCharStyleDropDown.selectedIndex + "@,@" + myCharStyleCheckBox.checkedState + "@,@" + myRadio1.checkedState + "@,@" + myRadio2.checkedState + "@,@" + myRadio3.checkedState + "@,@" + myRadio4.checkedState + "@,@" + myRoundto + "@,@" + myCustomGrepCheckBox.checkedState + "@,@" + myGrepField.editContents + "@,@" + myRadioStory.checkedState + "@,@";
				myTempFile.open( 'w', 'TEXT', 'R*ch');
				myTempFile.writeln(myTempFileContents);
				myTempFile.close();
				}
			catch(e){alert(e);}
		}//end of if dialog is not canceled
//End of numberAdjuster.jsx