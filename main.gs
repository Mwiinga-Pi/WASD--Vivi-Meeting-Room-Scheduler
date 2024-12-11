// INCLUDE SERVICES:
//      Calendar, Slides
// this connects to Second Trial of Things
// presentation/d/145ZJ4Rk1Hny9Ea44LurngFSC5A436o7kfvys_qV5Tt4/edit#slide=id.p
function fullCreateEventNotice() {
  var calendarId = '35p8q68b7vt7brl1h4g022emv6um04fn@import.calendar.google.com';  // Use 'primary' for the primary calendar or replace with your calendar ID
  var now = new Date();
  var oneDayFromNow = new Date(now.getTime() + 1 *  24 * 60 * 60 * 1000);
  

  var events = Calendar.Events.list(calendarId, {
    timeMin: now.toISOString(),
    timeMax: oneDayFromNow.toISOString(),
    showDeleted: false,
    singleEvents: true,
    orderBy: "startTime"
  }).items;

  //removeing the following line because we will use an existing Slides presentation (145ZJ4Rk1Hny9Ea44LurngFSC5A436o7kfvys_qV5Tt4)
  // var presentationId = createGoogleSlidesPresentation();
  var presentationId = '145ZJ4Rk1Hny9Ea44LurngFSC5A436o7kfvys_qV5Tt4'
  var presentation = SlidesApp.openById(presentationId);
  // console.log(presentationId) //this will show us the Slides presentationID
  // This is the link of a created Slides doc: https://docs.google.com/presentation/d/145ZJ4Rk1Hny9Ea44LurngFSC5A436o7kfvys_qV5Tt4/edit#slide=id.p
  // This is the PresentationID for said slide: 145ZJ4Rk1Hny9Ea44LurngFSC5A436o7kfvys_qV5Tt4

  var colEventMax1 = 7 //THIS IS USED TO DETERMIN AT WHAT POINT TO STOP LISTING EVENTS FOR A SINGLE COLUMN
  var eventArr = []
  var startArr = [] //create array for start of Title Lines
  var start2ndArr =[]
  var endArr = []  //create array for start of Title Lines
  var end2ndArr = []
  var stringIndex = 0           // start counter for index of the current index of the full string
  var arrayIndexer = 0; 
  var lineCount = 0
  var count = 0
  var fullEventList = "———————————————————\n"
  events.forEach(function(event){
    const time_start = event.start.dateTime.slice(11,16)
    const time_end = event.end.dateTime.slice(11,16)
    // console.log(time_end)
    var eventLoca = "";

    //this if else block just picks out descriptions that have unneccesary 
    if (event.description.includes("| x")){
      // console.log("Hello there")
      eventLoca = event.description.split("| x").pop();
    }  else if (event.description.includes(" | ")){
      // console.log("Hello there")
      eventLoca = event.description.split(" | ").pop();
    }  else if (event.description.includes("x")){
      // console.log("Hello there")
      eventLoca = event.description.split("x").pop();
    } else {
      eventLoca = event.description;
    }
    const des_text = event.summary + "\n    Time - " + time_start + " - " + time_end + "\n    Location - " + eventLoca + "———————————————————";
    count = eventArr.push(event);
    fullEventList = fullEventList + event.summary.toLocaleUpperCase('en-US') + "\n    Time - " + time_start + /*" - " + time_end + "\n    Location - "*/ "   |   " + eventLoca + "\n———————————————————\n"
    
    
    
  })
    
    //console.log(fullEventList)
    //console.log("The number of lines for this date:  ", fullEventList.split(/\r\n|\r|\n/).length)

    const iterator = fullEventList[Symbol.iterator]();
    var theChar = iterator.next();

    // var stringIndex = 0           // start counter for index of the current index of the full string
    // var arrayIndexer = 0; 
    // var startArr = [] //create array for start of Title Lines
    // var endArr = []  //create array for start of Title Lines
    // var lineCount = 0 
    while (!theChar.done) {      // while not the end of the string...
      // console.log("The current letter is:  ", theChar.value); // print the current char
      // console.log("the index of this letter is:  ", stringIndex)    // print the current index
     
      
      
      if (theChar.value == '\n'){
        lineCount += 1;
        if (lineCount % 3 === 1 ){
          let tempNum = stringIndex +1
          startArr.push(tempNum)                            // This is getting the start of the correct rows currently
          endArr.push(fullEventList.indexOf('\n', tempNum)) // This is getting the end of the correct rows currently
          // console.log("The current letter is:  ", theChar.value, " And the index of this letter is:  ", stringIndex); // This is getting the start of the correct rows currently
          // console.log(`The character at index ${tempNum} is ${fullEventList.charAt(tempNum)}`);
          // console.log("The end index of this line is:  ", fullEventList.indexOf('\n', tempNum)); // This is getting the end of the correct rows currently
          //console.log(arrayIndexer)
          arrayIndexer += 1
        }
        
        // console.log(arrayIndexer)
      }
      
      theChar = iterator.next();  //
      stringIndex ++
    }
    startArr.pop();
    endArr.pop();
    console.log(startArr)
    console.log(endArr)

    /**************************************************/
    //THIS IS HOW I GET THE INDEXS FOR THE EVENT TITLES FOR A 1 COLLUMN SLIDE
    /*for (let incri = 0; incri < 5; incri-=-1){
      console.log(startArr[incri] + " until " + endArr[incri])
    }*/
    /**************************************************/

  console.log("The number of events is: " + count + " We recorded " + fullEventList.split(/\r\n|\r|\n/).length + " lines")
    if (count <colEventMax1){
      console.log("function for 1 column")
      createSlide1Col(presentationId, fullEventList, startArr, endArr);
      // NEED TO FIND OUT WHERE THE END OF THE 5TH EVENT IS AND 'POP' THE REST OF THE ARRAY OFF TO NOT BE USED
    }
    else {
      console.log("function for 2 columns")
      createSlide2Col(presentationId, fullEventList, startArr, endArr);
      // NEED TO FIND OUT WHERE THE END OF THE 5TH EVENT IS AND 'POP' THE REST OF THE ARRAY INTO ANOTHER ARRAY (COL2EVENTS)
    }
  //createSlide(presentationId, fullEventList, startArr, endArr);
    //addTextBox(presentationId,Utilities.getUuid(),event)
  
  Logger.log('Presentation created with ID: ' + presentationId);
}

/**************************************************************************/
/*                            1 column function                           */
/**************************************************************************/
function createSlide1Col(presentationId, fullEventList, startArr, endArr) {
  deleteAllSlides()
  // You can specify the ID to use for the slide, as long as it's unique.
  const pageId = Utilities.getUuid();  //creates quasi-unique ID to use

  /*Getting and setting the current date*/
  var today = new Date();
  
  // Format the date in the style of "Saturday, August 3"
  var options = { weekday: 'long', month: 'long', day: 'numeric' };
  var formattedDate = today.toLocaleDateString('en-US', options);
  //console.log(formattedDate)
  /*Ending the setting the current date*/

  const requests = [{
    'createSlide': {
      'objectId': pageId,
      'insertionIndex': 0, 
      'slideLayoutReference': {
        'predefinedLayout': 'BLANK'
      }
    }
  }];
  /*---------------------------------------------------*/
  const elementId = Utilities.getUuid();
  const elementIdDate = Utilities.getUuid();
  
  const pt350 = {
    magnitude: 350,
    unit: 'PT'
  };
  const pt375 = {
    magnitude: 375,
    unit: 'PT'
  };
  const pt60 = {
    magnitude: 60,
    unit: 'PT'
  };
  const pt320 = {
    magnitude: 320,
    unit: 'PT'
  };
  

  const text_requests = [
    {
      createShape: {
        objectId: elementId,
        shapeType: 'TEXT_BOX',
        elementProperties: {
          pageObjectId: pageId,
          size: {
            height: pt350,
            width: pt375
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: 190,
            translateY: 0,
            unit: 'PT'
          }
          
        }
      }
    },
    // Insert text into the box, using the supplied element ID.
    {
      insertText: {
        objectId: elementId,
        insertionIndex: 0,
        // @ts-ignore
        text: fullEventList, //des_text
      }
    },
    {
      updateTextStyle: {
        objectId: elementId,
        textRange: {
          type: 'ALL',
        },
        style: {
          //fontFamily: 'Times New Roman',
          fontSize: {
            magnitude: 14,
            unit: 'PT'
          },
        },
        
        fields: 'fontSize'
      }
    }
    /* The following will get appended for the number of events that there is:
      {
          updateTextStyle: {
            objectId: elementId,
            textRange: {
              type: 'FIXED_RANGE',
              startIndex: startArr[value],
              endIndex: endArr[value]
            },
            style: {
              bold: true,
              italic: true
            },
            fields: 'bold,italic'
          }*/
  ];

  // Setting Date Request
  const date_requests = [
    {
      createShape: {
        objectId: elementIdDate,
        shapeType: 'TEXT_BOX',
        elementProperties: {
          pageObjectId: pageId,
          size: {
            height: pt60,
            width: pt320
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: 350,
            translateY: 340,
            unit: 'PT'
          }
          
        }
      }
    },
    // Insert text into the box, using the supplied element ID.
    {
      insertText: {
        objectId: elementIdDate,
        insertionIndex: 0,
        // @ts-ignore
        text: formattedDate, //des_text
      }
    },
    {
      updateTextStyle: {
        objectId: elementIdDate,
        textRange: {
          type: 'ALL',
        },
        style: {
          //fontFamily: 'Times New Roman',
          fontSize: {
            magnitude: 25,
            unit: 'PT'
          },
          foregroundColor: {
          opaqueColor: {
            rgbColor: {
              blue: 0.0,
              green: 1.0,
              red: 1.0
            }
          }
        },
        bold: true,
        },
        
        fields: 'foregroundColor, fontSize, bold'
      }
    }, 
    {
    updateParagraphStyle: {
      objectId: elementIdDate,
      style: {
        alignment: 'END'
      },
      fields: 'alignment'
    }}
  ];



  try {
    const slide =
      Slides.Presentations.batchUpdate({'requests': requests}, presentationId);
      const newSlideID = slide.replies[0].createSlide.objectId
    console.log('Created Slide with ID: ' + newSlideID);
    //addTextBox(presentationId, newSlideID, event);
    
    const iterator1 = startArr[Symbol.iterator]();
    for (let value = 0; value < startArr.length;value++){
    text_requests.push(
        {
          updateTextStyle: {
            objectId: elementId,
            textRange: {
              type: 'FIXED_RANGE',
              startIndex: startArr[value],
              endIndex: endArr[value]
            },
            style: {
              bold: true,
              italic: true
            },
            fields: 'bold,italic'
          }
        }
    )}
    const createTextboxWithTextResponse = Slides.Presentations.batchUpdate({requests: text_requests}, presentationId);
    //const createShapeResponse = createTextboxWithTextResponse.replies[0].createShape;
    const createTextbox3WithDate = Slides.Presentations.batchUpdate({ requests: date_requests}, presentationId)
    return slide;
  } catch (e) {
    // TODO (developer) - Handle Exception
    console.log('Failed with error %s', e.message);
  }
}

function deleteAllSlides() {
  // ID of the Google Slides presentation
  var slidesId = '145ZJ4Rk1Hny9Ea44LurngFSC5A436o7kfvys_qV5Tt4';

  // Retrieve the list of slides
  var presentation = SlidesApp.openById(slidesId);
  var slides = presentation.getSlides();
  
  // Create delete requests for each slide
  var requests = slides.map(slide => ({
    deleteObject: {
      objectId: slide.getObjectId()
    }
  }));
  
  // Execute the batchUpdate request to delete all slides
  Slides.Presentations.batchUpdate({requests: requests}, slidesId);
  
  Logger.log('All slides deleted.');
}
/**************************************************************************/
/*                            2 column function                           */
/**************************************************************************/
function createSlide2Col(presentationId, fullEventList, startArr, endArr) {
  deleteAllSlides()
  var start2ndArr =[]
  var end2ndArr =[]
  var stringIndex = 0;
  var lineCount = 0;
  var arrayIndexer = 0; 
  // You can specify the ID to use for the slide, as long as it's unique.
  var firstColofEvent = fullEventList.slice(0,startArr[6])
  let startIndexof2ndHalfofEvents = startArr[6] ;
  var secondColofEvent = "——————————————————\n" + fullEventList.slice(startIndexof2ndHalfofEvents)
  //console.log(firstColofEvent) //verifies that the first column pulls the correct set of events
  //console.log(secondColofEvent) //verifies that the second column pulls the correct set of events
  const pageId = Utilities.getUuid();  //creates quasi-unique ID to use

  /*Getting and setting the current date*/
  var today = new Date();
  
  // Format the date in the style of "Saturday, August 3"
  var options = { weekday: 'long', month: 'long', day: 'numeric' };
  var formattedDate = today.toLocaleDateString('en-US', options);
  //console.log(formattedDate)
  /*Ending the setting the current date*/

  const iterator2o2 = secondColofEvent[Symbol.iterator]();
    var theChar = iterator2o2.next();

    while (!theChar.done) 
    {      // while not the end of the string...
      
      if (theChar.value == '\n'){
        lineCount += 1;
        if (lineCount % 3 === 1 ){
          let tempNum = stringIndex +1
          start2ndArr.push(tempNum)                            // This is getting the start of the correct rows currently
          end2ndArr.push(secondColofEvent.indexOf('\n', tempNum)) // This is getting the end of the correct rows currently
          arrayIndexer += 1
        }
        
        // console.log(arrayIndexer)
      }
      
      theChar = iterator2o2.next();  //
      stringIndex ++
    }
    start2ndArr.pop();
    end2ndArr.pop();
    console.log("2nd column arrays")
    console.log(start2ndArr)
    console.log(end2ndArr)


  const requests = [{
    'createSlide': {
      'objectId': pageId,
      'insertionIndex': 0, 
      'slideLayoutReference': {
        'predefinedLayout': 'BLANK'
      }
    }
  }];
  /*---------------------------------------------------*/
  const elementIdCol1of2 = Utilities.getUuid();
  const elementIdCol2of2 = Utilities.getUuid();
  const elementIdDate = Utilities.getUuid();
  const pt60 = {
    magnitude: 60,
    unit: 'PT'
  };
  const pt320 = {
    magnitude: 320,
    unit: 'PT'
  };
  const pt350 = {
    magnitude: 350,
    unit: 'PT'
  };
  const pt375 = {
    magnitude: 375,
    unit: 'PT'
  };

  

  const text_requestsCol1 = [
    {
      createShape: {
        objectId: elementIdCol1of2,
        shapeType: 'TEXT_BOX',
        elementProperties: {
          pageObjectId: pageId,
          size: {
            height: pt350,
            width: pt375
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: 10,
            translateY: 0,
            unit: 'PT'
          }
          
        }
      }
    },
    // Insert text into the FIRST COLUMN, using the supplied element ID (elementIdCol1of2).
    {
      insertText: {
        objectId: elementIdCol1of2,
        insertionIndex: 0,
        // @ts-ignore
        text: firstColofEvent, //des_text
      }
    },
    {
      updateTextStyle: {
        objectId: elementIdCol1of2,
        textRange: {
          type: 'ALL',
        },
        style: {
          //fontFamily: 'Times New Roman',
          fontSize: {
            magnitude: 14,
            unit: 'PT'
          },
        },
        
        fields: 'fontSize'
      }
    }
    /* The following will get appended for the number of events that there is:
      {
          updateTextStyle: {
            objectId: elementId,
            textRange: {
              type: 'FIXED_RANGE',
              startIndex: startArr[value],
              endIndex: endArr[value]
            },
            style: {
              bold: true,
              italic: true
            },
            fields: 'bold,italic'
          }*/
  ];

  const text_requestsCol2 = [
    {
      createShape: {
        objectId: elementIdCol2of2,
        shapeType: 'TEXT_BOX',
        elementProperties: {
          pageObjectId: pageId,
          size: {
            height: pt350,
            width: pt375
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: 350,
            translateY: 0,
            unit: 'PT'
          }
          
        }
      }
    },
    // Insert text into the FIRST COLUMN, using the supplied element ID (elementIdCol2of2).
    {
      insertText: {
        objectId: elementIdCol2of2,
        insertionIndex: 0,
        // @ts-ignore
        text: secondColofEvent, //des_text
      }
    },
    {
      updateTextStyle: {
        objectId: elementIdCol2of2,
        textRange: {
          type: 'ALL',
        },
        style: {
          //fontFamily: 'Times New Roman',
          fontSize: {
            magnitude: 14,
            unit: 'PT'
          },
        },
        
        fields: 'fontSize'
      }
    }
  ];

  

  // Setting Date Request
  const date_requests = [
    {
      createShape: {
        objectId: elementIdDate,
        shapeType: 'TEXT_BOX',
        elementProperties: {
          pageObjectId: pageId,
          size: {
            height: pt60,
            width: pt320
          },
          transform: {
            scaleX: 1,
            scaleY: 1,
            translateX: 350,
            translateY: 340,
            unit: 'PT'
          }
          
        }
      }
    },
    // Insert text into the box, using the supplied element ID.
    {
      insertText: {
        objectId: elementIdDate,
        insertionIndex: 0,
        // @ts-ignore
        text: formattedDate, //des_text
      }
    },
    {
      updateTextStyle: {
        objectId: elementIdDate,
        textRange: {
          type: 'ALL',
        },
        style: {
          //fontFamily: 'Times New Roman',
          fontSize: {
            magnitude: 25,
            unit: 'PT'
          },
          foregroundColor: {
          opaqueColor: {
            rgbColor: {
              blue: 0.0,
              green: 1.0,
              red: 1.0
            }
          }
        },
        bold: true,
        },
        
        fields: 'foregroundColor, fontSize, bold'
      }
    }, 
    {
    updateParagraphStyle: {
      objectId: elementIdDate,
      style: {
        alignment: 'END'
      },
      fields: 'alignment'
    }}
  ];


  try {
    const slide =
      Slides.Presentations.batchUpdate({'requests': requests}, presentationId);
      const newSlideID = slide.replies[0].createSlide.objectId
    console.log('Created Slide with ID: ' + newSlideID);
    //addTextBox(presentationId, newSlideID, event);
    
    const iterator1 = startArr[Symbol.iterator]();
    for (let value = 0; value < 6/*startArr.length*/;value++){
    text_requestsCol1.push(
        {
          updateTextStyle: {
            objectId: elementIdCol1of2,
            textRange: {
              type: 'FIXED_RANGE',
              startIndex: startArr[value],
              endIndex: endArr[value]
            },
            style: {
              bold: true,
              italic: true
            },
            fields: 'bold,italic'
          }
        }
    )}
    for (let value = 0; value < start2ndArr.length;value++){
    text_requestsCol2.push(
        {
          updateTextStyle: {
            objectId: elementIdCol2of2,
            textRange: {
              type: 'FIXED_RANGE',
              startIndex: start2ndArr[value],
              endIndex: end2ndArr[value]
            },
            style: {
              bold: true,
              italic: true
            },
            fields: 'bold,italic'
          }
        }
    )}
    const createTextbox1WithTextResponse = Slides.Presentations.batchUpdate({requests: text_requestsCol1/*, requests: text_requestsCol2*/}, presentationId);
      Slides.Presentations.batchUpdate({ requests: text_requestsCol2}, presentationId);
    const createTextbox3WithDate = Slides.Presentations.batchUpdate({ requests: date_requests}, presentationId)
    //const createShapeResponse = createTextboxWithTextResponse.replies[0].createShape;
    
    var today = new Date();
  
  // Format the date in the style of "Saturday, August 3"
  var options = { weekday: 'long', month: 'long', day: 'numeric' };
  var formattedDate = today.toLocaleDateString('en-US', options);
  console.log(formattedDate)

    return slide;
  } catch (e) {
    // TODO (developer) - Handle Exception
    console.log('Failed with error %s', e.message);
  }
}


//https://developers.google.com/slides/api/guides/add-shape?authuser=2#apps-script
// https://developers.google.com/calendar/api/v3/reference/events?authuser=2
