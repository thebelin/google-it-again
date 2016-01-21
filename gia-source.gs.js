/**
 * ------------------------------------------------------------------
 * Library Functions for googling it again
 * ------------------------------------------------------------------
 *
 * Project Key: McpfOjy0GA2kDM2ytVgOhYwIv2K89z2yK
 *
 * This set of tools is designed to make it easier to build web applications from google
 * docs spreadsheet data. Since they are something used over and over again, they are called
 * Google It Again, or G.I.A.
 * 
 * @author Belin Fieldson <thebelin@gmail.com>
 */

/**
 * @type string Holds the localId property (which should be set in the embedding script)
 */
var localId = '',

  /**
   * Holds a collection of sheets which aren't returned in a getAllData function
   * This can be overridden by the calling script as an attribute
   *
   * @type array
   */
  protectedSheets = ['users', 'roles'],

  /**
   * Will hold all the routes (which should be set in the embedding script)
   * an example version of the 'none' route is included which gets all the data in all the sheets
   * an example of the 'poll' route is also included which only returns the data if the hash is different
   * and returns it on the public route
   *
   * @type object
   */
  router = {

    'none': function (e) {
      // return all the data in the sheets not in the public data adapter
      return serveJSONP(e, getAllData());
    },

    'poll': function (e) {
      var allData = getAllData();
      if (e.parameters.hash[0] !== allData.hash) {
        return serveJSONP(e, allData);
      }
      return serveJSONP(e, allData.hash);
    }
  },

  /**
   * A local cache object for function re-use
   *
   * @type object
   */
  cache = {
    // The local cache data
    d : {},

    // Set a new item
    set : function (key, value) {
      this.d[key] = value;
      return this.d[key];
    },

    // get a cache value (or set and return default while getting)
    get : function (key, defaultVal) {
      return this.d[key] || defaultVal ? this.set(key, defaultVal) : false;
    },

    // Make a cache identifier with an MD5 signature
    makeKey : function (key, args) {
      return GetMD5Hash(key + JSON.stringify(args));
    },

    // Clear all the cached data
    clear : function () {
      this.d = {};
    }
  };

/**
 * The Main route request, runs each web instance
 *
 * @param {object} e The request event descriptor
 * @return {object} HTTP output
 */
function doGet(e) {
  if (e instanceof Object) {
    /**
    * Decide how to run the router based on the a parameter
    */
    if (e.parameters.action) {
      // This calls the function asked for in the "action" parameter, if it exists
      if (router[e.parameters.action] instanceof Function) {
        return router[e.parameters.action].apply(this, [e]);
      }
    } else {
      // This calls the default function
      if (router.none && router.none instanceof Function) {
        return router.none.apply(this, [e]);
      }
    }
  }
}

/**
* A prototype for a simple object
*
* @param {array} keys   The object keys
* @param {array} values The values to put in the keys
* @param {int}   lineId {optional} The line identifying this object
*/
function GenericObject(keys, values, lineId) {
  var i;
  // If there's an index (lineId) then record it as _id
  if (lineId) {
    this._id = parseInt(lineId, 10);
  }
  // Iterate the keys
  for (i in keys) {
    if (keys.hasOwnProperty(i)) {
      // If there's a corresponding data value, apply it to this
      if (values[i] !== 'undefined') {
        this[keys[i]] = values[i];
      }
    }
  }
}

/**
 * Generic object converter
 * Converts an array of arrays to an array of objects
 *
 * @param {array} arr The array of arrays to convert, first item should be keys
 *
 * @return {object} key-value set of the same data
 */
function convertToObject(arr) {
  // @var {array} ret The return value
  var ret = [],

  // @var {array} objKeys Get the first row, those are the object keys
    objKeys = arr[0] instanceof Array ? arr[0] : [],

  // @type {string} i Key iterator
    i,

  // @type {object} thisObject an Object to hold array values
    thisObject;

  // Convert each successive row and add to ret
  for (i = 0; i < arr.length; i++) {
    if (i !== 0) {
      thisObject = new GenericObject(objKeys, parseContents(arr[i]), i);
      ret.push(thisObject);
    }
  }
  return ret;
}

/**
 * A safe parse function which either returns parsed results or the value passed in
 *
 * @param {string} val A string possibly containing convertable data.
 */
function parseIt(val) {
  var ret;
  try {
    ret = JSON.parse(val);
  } catch (e) {
    ret = val;
  }
  return ret;
}

/**
 * Parse each item in the param, convert it if it's a JSON value
 */
function parseContents(arr) {
  var ret = [],
    attr;
  for (attr in arr) {
    if (arr.hasOwnProperty(attr)) {
      // Parse each attribute via JSON, if it doesn't parse, then set it to what it is
      ret.push(parseIt(arr[attr]));
    }
  }
  return ret;
}

/**
 * Get specified sheet
 *
 * @param name The name of the sheet
 *
 * @return A apreadsheet app object, or null
 */
function getSheetValues(name) {
  var ret = cache.get(name);
  if (ret) {
    return ret;
  }

  var vals = SpreadsheetApp.openById(localId).getSheetByName(name).getDataRange().getValues();
  return cache.set(name, convertToObject(vals));
}

/**
 * Get specified sheet headings as array of strings
 * 
 * @param name The name of the sheet
 *
 * @return array
 */
function getSheetHeaders(name) {
  var ret = cache.get('heading' + name),
    vals = getSheetValues(name);

  if (ret) {
    return ret;
  }

  // Return the first object's keys
  return cache.set('heading' + name, Object.keys(vals[0]));
}

/**
 * Save a row in the specified sheet
 * 
 * @param sheet The Sheet which will hold the new row
 * @param data  The data to write to the sheet
 */
function saveRow(sheet, data) {
  if (!data._id) {
    // don't save unless there's an _id set
    return false;
  }
  // increment the _id value to account for the header row
  data._id++;

  // @var array The headers as an array
  var headers = getSheetHeaders(sheet),
    // The range to write to according to the _id attribute of the data
    // The _id attribute is set by the read routine and corresponds to the row
    writeRange = SpreadsheetApp.openById(localId)
      .getSheetByName(sheet)
      .getRange(data._id, 1, 1, headers.length - 1),

    // The actual writable range
    writeValues = writeRange.getValues(),

    // If there's a hash field, this is it
    hashId,

    // an iterator
    i;

  // Iterate the headers and set them on the write values
  for (i in headers) {
    if (headers.hasOwnProperty(i)) {
      if (data[headers[i]] !== 'undefined') {
        if (headers[i] === '_hash') {
          writeValues[0][i - 1] = '';
          hashId = i - 1;
        } else if (headers[i] !== '_id') {
          writeValues[0][i - 1] = data[headers[i]];
        }
      } else {
        // The input value is undefined, set write value to null
        writeValues[0][i - 1] = null;
      }
    }
  }

  // update the signature hash if it exists
  if (typeof hashId === 'number') {
    writeValues[0][hashId] = GetMD5Hash(JSON.stringify(writeValues[0]));
  }
  writeRange.setValues(writeValues);
}

/**
 * Create a new row in the target sheet filled with the specified data
 * return the data, with the _id inserted
 * 
 * @param sheet The Sheet which willhold the new row
 * @param data  The data to write to the sheet
 */
function createRow(sheet, data) {
  var vals = getSheetValues(sheet),
    ret = data;

  // Iterate the id of the returned data by one
  ret._id = ++vals.length;
  saveRow(sheet, ret);
  return ret;
}

/**
 * There might be post visitors too, though they shouldn't be cross domain posting.
 */
function doPost(e) {
  doGet(e);
}

/**
 * Get a sheet as an object by converting all the first headings to the keys
 *
 * @param string sheetName The name of ths google doc to convert
 * @return object key-value object with the data
 */
function getSheet(sheetName) {
  var ret = cache.get(sheetName);
  if (ret) {
    return ret;
  }
  return cache.set(sheetName, getSheetValues(sheetName));
}

/**
 * Remove the protected sheets from the array of sheets
 *
 * @param array sheets          Sheet objects from SpreadsheetApp getSheets()
 * @param array protectedSheets An array of strings which define which sheets are protected
 *
 * @return array The sheets minus any protected ones
 */
function removeProtected(sheets, protectedSheets) {
  // The data to return
  var returnSheets = [],

    // Whether or not to do the add to the return data set
    doAdd = true;

  if (sheets.length) {
    sheets.map(function (sheet) {
      protectedSheets.map(function (protectedSheet) {
        if (sheet.getName() === protectedSheet) {
          doAdd = false;
        }
      });
      // This is an OK sheet, add it to the output
      if (doAdd) {
        returnSheets.push(sheet);
      }
    });
  }
  return returnSheets;
}
/**
 * Get the user data for the specified parameters (userid / userkey)
 *
 * @param Object e The data from the request
 *
 * @return Array User information about the user data posted, or empty array
 */
function getApiUserData(e) {
  return getSheetValues('apiusers').map(function (thisUser) {
    if (thisUser.apiUser === e.parameters.userid[0] && thisUser.apiKey === e.parameters.userkey[0]) {
      // @todo: write the user record indicating that the user just accessed the API
      return thisUser;
    }
  });
}

/**
 * Handle API routing
 *
 * @param e object Get event object
 * @param sheetName string The name of the sheet to route to
 */
function doApiRoute(e, sheetName) {
  // Check the user data posted
  var userData = getApiUserData(e),

    // Get the api route to do
    apiRoute = e.parameters.method || 'GET',

    // Create all the routes to use for this sheet
    secureRoutes = makeEndpoint(sheetName);

  if (userData.length && userData[0]) {
    // This request is allowed into user protected routes
    if (typeof secureRoutes[apiRoute] === 'function') {
      // return the specified executed route, arguing the event and the userData
      return serveJSON(e, secureRoutes[apiRoute](e));
    }
  } else {
    return serveJSON(e, {error: "user parameters for API access are incorrect"});
  }

  // This user is requesting an invalid resource
  Logger.log("User requested invalid resource: " + JSON.stringify(e));
}

/**
 * Create a secure (private) endpoint representing the specified sheet
 * secured based on data in the local apiusers sheet
 *
 * @param string sheetName The name of the sheet to for the endpoint being created
 *
 * @return object a router for use as a gia.router
 */
function makeEndpoint(sheetName) {
  // Get the properties for the intended object
  var paramArray = getSheetHeaders(sheetName),

    /**
     * Get the parameters in the request object for this API access
     * 
     * @param object e          Google Apps Script Request
     * @param array  paramArray The values to use in the API operation (read/write or to filter by)
     * 
     * @return object An ApiParams static object
     */
    getApiParams = function (e, paramArray) {
      return {
        // Default the page to 0, but if they provided a page number, subtract one from it, minimum of 0
        'page': (e.parameters.page && e.parameters.page[0] && !isNaN(e.parameters.page[0] - 1)) ? Math.min(parseInt(e.parameters.page[0] - 1, 10), 0) : 0,

        // Default the records per page limit to 10, but no higher than 250
        'limit': (e.parameters.limit && e.parameters.limit[0] && !isNaN(e.parameters.limit) && parseInt(e.parameters.limit[0], 10) <= 250) ? parseInt(e.parameters.limit[0], 10) : 10,

        // filter parameters should return an object with a key for each filtered parameter
        'filters': (function () {
          // to be returned from filter map
          var retObj = {};
          paramArray.map(function (param) {

            // Set the filters in the map
            if (e.parameters.hasOwnProperty(param)) {
              retObj[param] = e.parameters[param][0];
            } else {
              retObj[param] = null;
            }
          });
          return retObj;
        }())
      };
    },

    /**
     * filter the sheet values according to the filters array
     *
     * @param object e           Google Apps Script Request
     * @param object sheetValues All the values from a sheet
     * @param array  filters     The url parameters which can be used as filters
     * 
     * @return object The sheetValues, but filtered
     */
    filterData = function (e, sheetValues, filters) {
      var outValues = [];
      // Get record(s) according to the apiParams.filters property
      if (filters && filters.constructor === Array) {
        // when applying map to the sheetValues array, it returns nulls for each if returned from the map directly
        sheetValues.map(function (record) {
          var compare = true, i;
          for (i in filters) {
            if (e.parameters[i] && e.parameters[i][0] !== record[i]) {
              compare = false;
            }
          }
          // only add valid items to outValues
          if (compare) {
            outValues.push(record);
          }
        });
      }
      return outValues;
    };


  // This is the router to return which provides the api through gia.router
  return {
    /**
     * Read route: use to get a user and display them
     */
    'GET': function (e) {
      // Set the apiParams according to paramArray
      var apiParams = getApiParams(e, paramArray),

      // Get the sheet to be returned and filter it
        sheetValues = filterData(e, getSheetValues(sheetName), extend({enabled: true}, apiParams.filters));

      // return the limited result set:
      return sheetValues.slice(apiParams.page * apiParams.limit, apiParams.limit);
    },

    /**
     * Create route: Use to create a new and return the values
     */
    'PUT': function (e) {
      // Set the apiParams according to paramArray
      var apiParams = getApiParams(e, paramArray),

      // Insert the record into the sheet, according to data in the filters
        rowData = createRow(sheetName, extend({enabled: true}, apiParams.filters));

      // Return the created object, with the _id value
      return rowData;

    },

    /**
     * Delete Route - use to disable
     */
    'DELETE': function (e) {
      // Set the apiParams according to paramArray
      var apiParams = getApiParams(e, paramArray);

      // Force the apiParams _id property to 0 or whatever it is
      apiParams._id = (apiParams.filters._id || 0);

      // update the specified record _id with an enabled = false
      saveRow(sheetName, {enabled: false, _id: apiParams._id});

      // Return the updated object
      return {enabled: false, _id: apiParams._id};
    },

    /**
     * POST Route - use to update
     */
    'POST': function (e) {
      // Set the apiParams according to paramArray
      var apiParams = getApiParams(e, paramArray),

      // Get the sheet to be added to and filter it by the specified _id in the POST
      // Force the apiParams _id property to 0 or whatever it is
        sheetValues = filterData(e, getSheetValues('users'), {_id: (apiParams.filters._id || 0), enabled: true }),
        newValues = extend(sheetValues, apiParams.filters);

      // Insert the record into the sheet, according to data in the filters
      saveRow(sheetName, newValues);

      // Return the updated object
      return newValues;
    }
  };
}

/**
 * Get all the values out of all the unprotected sheets
 * 
 * @param array sheets (optional) An array of sheet objects from SpreadsheetApp getSheets()
 *
 * @return object A json object to send to the user with all the data
 */
function getAllData(sheets) {
  // @type object The return data
  var returnData = {};
  // Set the argued sheets to either themselves or the entire collection of sheets minus protected ones,
  // if nothing was specified.
  sheets = sheets instanceof Array ? sheets : removeProtected(SpreadsheetApp.openById(localId).getSheets(), protectedSheets);
  if (sheets.length) {
    sheets.map(function (sheet) {
      returnData[sheet.getName()] = getSheetValues(sheet.getName());
    });
  }
  // Give a signature which changes when the data does
  returnData.hash = GetMD5Hash(JSON.stringify(returnData));
  return returnData;
}

/**
 * Output in response to a JSONP request
 *
 * @param e       object Event Context
 * @param content mixed  Values to be displayed
 * 
 * @return object Content ready for app context
 */
function serveJSONP(e, content) {
  if (e.parameters.prefix && e.parameters.prefix !== 'undefined') {
    return ContentService.createTextOutput(
      e.parameters.prefix + '(' + JSON.stringify(content) + ')'
    ).setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  // no prefix, send JSON instead
  return serveJSON(e, content);
}

/**
 * Output in response to a JSON request
 *
 * @param e       object Event Context
 * @param content mixed  Values to be displayed
 * 
 * @return object Content ready for app context
 */
function serveJSON(e, content) {
  return ContentService.createTextOutput(
    JSON.stringify(content)
  ).setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/**
 * Convert a column heading number to letter
 *
 * Column A is 1
 * Column AA is 27
 * Column BA is 53...
 */
function convertToColumn(numIndex) {
  // A = 65
  var colGroup = Math.floor(numIndex / 26),
    ord = numIndex % 26;

  return colGroup ? String.fromCharCode(64 + colGroup) + String.fromCharCode(64 + ord)
    : String.fromCharCode(64 + ord);
}

/**
 * Get only the values preceding non null of the passed in array
 */
function getNotNull(values) {
  var ret = [],
    i;
  for (i = 0; i <= values.length; i++) {
    if (values[i] === '' || values[i] === null || values[i][0] === '' || values[i][0] === null) {
      break;
    }
    ret.push(values[i]);
  }
  return ret;
}

// From Stack Overflow:
function GetMD5Hash(input) {
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, input),
    txtHash = '',
    hashVal,
    j;

  for (j = 0; j < rawHash.length; j++) {
    hashVal = rawHash[j];
    if (hashVal < 0) {
      hashVal += 256;
    }
    if (hashVal.toString(16).length === 1) {
      txtHash += "0";
    }
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}

/**
 * Add all the passed in key-values together
 * Overlay values passed in sequentially
 *
 * @param {object] data The data to combine
 * @param {object] data1 The data to combine
 * @param {object] data2 The data to combine
 * @param {object] data3 The data to combine
 * @param {object] data4 The data to combine
 * ...
 *
 * @return {object} The data added together
 */
function extend() {
  // @var object The return value should include the default getParams
  var output = typeof arguments[0] === 'object' ? arguments[0] : {},

    // @var int i Iterator for number of objects to combine
    i = 1,

    // @var string a Attribute to clone
    a = '';

  for (i = 1; i < arguments.length; i++) {
    // Add all the keys from the data
    if (typeof arguments[i] === 'object') {
      for (a in arguments[i]) {
        if (arguments[i].hasOwnProperty(a)) {
          output[a] = arguments[i][a];
        }
      }
    }
  }
  return output;
}

/**
 * Create a PDF from the specified template and fill the template in with
 * the values from tempVals object
 *
 * @param string templateId The google docs id of the template to use
 * @param string pdfName    The document filename to save the file as
 * @param object tempVals   An object containing keys for all the substitutions in the template
 * @param bool   truncate   set to true to truncate numbers to 2 digits
 *
 * @return {Object/bool} DriveApp.file or false
 */
function makePdfFromTemplate (templateId, pdfName, tempVals, truncate) {
  // Get the pdf, make a copy and then parse the contents into the copyBody
  var key, newFile, copyFile = DriveApp.getFileById(templateId).makeCopy(),
      copyId = copyFile.getId(),
      copyDoc = DocumentApp.openById(copyId),
      copyBody = copyDoc.getActiveSection();

  // Only process the doc if the settings have been properly provided
  if (typeof tempVals == 'object') {
    // Take the tempVals and place them in the copyBody where there are placeholders
    for (key in tempVals) {
      // If the value is a number, and truncation is enabled then truncate to 2 decimals max
      if (!truncate || isNaN(tempVals[key])) {
        copyBody.replaceText('%' + key + '%', tempVals[key]);
      } else {
        copyBody.replaceText('%' + key + '%', Math.round(tempVals[key] * 100) / 100 );
      }
    }
    
    // Save the temp document and close it
    copyDoc.saveAndClose();
     
    // Create the new pdf document
    newFile = DriveApp.createFile(copyFile.getAs('application/pdf'));
    if (pdfName !== '') {
      newFile.setName(pdfName);
    }
    
    // Trash the temp file which is a google doc file, not pdf
    copyFile.setTrashed(true);
    
    return newFile;
  }
  Logger.log('improper makePdf args: ' + typeof tempVals + typeof copyBody);
  return false;
}
