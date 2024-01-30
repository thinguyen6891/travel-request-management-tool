function processHotelAdd(formObject) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Hotel');
  var lastRow = sheet.getLastRow() + 1;
  sheet.appendRow([formObject.request_id,
                   'Booked',
                   formObject.name,
                   formObject.department,
                   getDate(formObject.checkin_date),
                   getDate(formObject.checkout_date),
                   formObject.location,
                   formObject.hotel_name,
                   '',
                   '',
                   formObject.price_per_night,
                   formObject.total_amount,
                   formObject.paymentSelect,
                   formObject.booking_code,
                   new Date().toLocaleDateString(),
                   formObject.booked_by
                   ]);
  sheet.getRange('I' + lastRow).insertCheckboxes().setValue(formObject.breakfast);
  sheet.getRange('J' + lastRow).setFormula('=DATEDIF(E' + lastRow + ',' + 'F' + lastRow + ',"D")');
}

function processTrainAdd(formObject) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Train');
  var lastRow = sheet.getLastRow() + 1;
  sheet.appendRow([formObject.request_id,
                   'Booked',
                   formObject.name,
                   formObject.department,
                   getDate(formObject.departure_date),
                   getDate(formObject.arrival_date),
                   formObject.departure_des,
                   formObject.arrival_des,
                   "",
                   formObject.carrier,
                   formObject.ticket_type,
                   formObject.total_amount,
                   formObject.paymentSelect,
                   formObject.booking_code,
                   new Date().toLocaleDateString(),
                   formObject.booked_by
                   ]);
  sheet.getRange('I' + lastRow).insertCheckboxes().setValue(formObject.return_check);
}

function processFlightAdd(formObject) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Flight');
  var lastRow = sheet.getLastRow() + 1;
  sheet.appendRow([formObject.request_id,
                   'Booked',
                   formObject.name,
                   formObject.department,
                   getDate(ormObject.departure_date),
                   getDate(formObject.arrival_date),
                   formObject.departure_des,
                   formObject.arrival_des,
                   '',
                   formObject.carrier,
                   formObject.ticket_type,
                   formObject.total_amount,
                   formObject.paymentSelect,
                   formObject.booking_code,
                   new Date().toLocaleDateString(),
                   formObject.booked_by
                   ]);
  sheet.getRange('I' + lastRow).insertCheckboxes().setValue(formObject.return_check);
}

function processRequestForm(formObject) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Requests');
  const id = sheet.getRange('C' + sheet.getLastRow()).getValue() + 1;
  const type_dict = {
    'Flight': getPlan(formObject.flight), 
    'Train': getPlan(formObject.train),
    'Hotel': getPlan(formObject.hotel)
  }

  const type = [type_dict['Flight'][0], type_dict['Train'][0], type_dict['Hotel'][0]].filter(Boolean).join(' + ');
  sheet.appendRow(['Not started',
                   '',
                   id,
                   type,
                   new Date(),
                   formObject.empEmail,
                   formObject.empName,
                   getDate(formObject.dob),
                   formObject.department,
                   formObject.travelReason,
                   formObject.supEmail,
                   type_dict['Flight'][1],
                   getDate(formObject.flightDepDate),
                   getDate(formObject.flightReDate),
                   formObject.flightDepLoc,
                   formObject.flightArrLoc,
                   formObject.flightSeat,
                   '',
                   '',
                   type_dict['Train'][1],
                   getDate(formObject.trainDepDate),
                   getDate(formObject.trainReDate),
                   formObject.trainDepLoc,
                   formObject.trainArrLoc,
                   formObject.trainSeat,
                   '',
                   '',
                   getDate(formObject.checkinDate),
                   getDate(formObject.checkoutDate),
                   formObject.hotelLoc,
                   '',
                   formObject.breakfast
                   ]);
};

function getSelect(list,firstRow) {
  var sh = SpreadsheetApp.getActive().getSheetByName(list);
  var len = sh.getLastRow() + 1 - firstRow;
  var shValues = sh.getRange(firstRow,1,len,1).getValues();
  var options = '';
  for (var i in shValues) {
    var cellValue = shValues[i][0];
    options += '<option value=\"' + cellValue + '\">' + cellValue + '</option>';
  };
  return options;
};

function getNameFromEmail(email) {
  var namePieces = email.replace(/@.+$/, '').split('.');
  for (i=0; i<namePieces.length; i++) {
    namePieces[i] = namePieces[i].replace(/^[a-z]/, function(match) {
        return match.toUpperCase();
    });
  }
  return namePieces.join(' ');
};

function getDate(date){
  if(date) {
    return Utilities.formatDate(new Date(date), 'GMT', 'MM/dd/yyyy');
  } else {
    return '';
  }
};

function getPlan(type, returnDate) {
  if(type) {
    if (returnDate) {
      return [type, 'One Way'];
    } else {
      return [type, 'Return'];
    }
  } else {
    return ['',''];
  }
};