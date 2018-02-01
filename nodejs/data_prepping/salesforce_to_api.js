var excel_file = require('excel');
// var convertTime = require('convert-time');
var fs = require('fs');

/*
clean csv of previous data
*/
fs.truncate('case_origin_owner_timestamp_csatscore.csv', 0, function() {
  console.log('successfully removed previous content from csv file')
})

/*
set column names for csv file
*/
fs.writeFile('case_origin_owner_timestamp_csatscore.csv', "caseNumber, caseOrign, caseOwner, startDate, endDate, csatScore \n",
	function (err) {
		if (err)
			throw err;
	});

excel_file('taccsatscores.xlsx', function(err, data) {
  if(err) throw err;
    /*
    iterate through the csat score excelsheet & convert date/time to iso 8601
    */
    for (row = 1; row < 32; row++) {
    // currently hard setting the max row, this is a workaround for now until edgecase timestamps from
    // salesforce data are correclty formatted from d/m/yy -> yyyy-mm-dd
      current_row_data = data[row]
      raw_date_time = current_row_data[3]
      //console.log(current_row_data)

      /*
      split date & time, format date to yyyy-mm-dd, subtract 10 minutes from 24hr time as start time.
      add time to formatted date and join date time in iso 8601 format

      IMPORTANT: Ask Manoj how far back we need to go (in minutes) to find the actual call time from
      the opened time stamp in excel
      */
      var split_date = raw_date_time.split(" ")
      var date = split_date[0]

      var split_date_at_slash = date.split("/")
      var rearranged_date = split_date_at_slash[2] + "-" + split_date_at_slash[1] + "-" + split_date_at_slash[0]

      var time = split_date[1]
      var split_time_at_colon = time.split(":")
      var time_hour = split_time_at_colon[0]
      var time_minutes = split_time_at_colon[1]

        if (time_minutes <= 10.0 && time_hour < 2.0) {
            var time_hour_start = 24
            time_hour_start = time_hour_start - time_hour
            var time_minutes_start = time_minutes - 10.0
        }
        else {
          time_hour_start = time_hour--
          time_minutes_start = time_minutes - 10.0
        }

      /*
      extract case#, case origin, case owner and store in csv file
      */
      var case_number = current_row_data[0] + ","
      fs.appendFile('case_origin_owner_timestamp_csatscore.csv', case_number, function(error) {
  		    if (error) throw error;
  		});
      var case_origin = current_row_data[1] + ","
      fs.appendFile('case_origin_owner_timestamp_csatscore.csv', case_origin, function(error) {
  		    if (error) throw error;
  		});
      var case_owner = current_row_data[2] + ","
      fs.appendFile('case_origin_owner_timestamp_csatscore.csv', case_owner, function(error) {
  		    if (error) throw error;
  		});

      var start_time = "T" + time_hour_start + ":" + time_minutes_start + ":00.000Z" + ",";
      var rearranged_start_date_time = rearranged_date + start_time
      console.log("call start time: " + rearranged_start_date_time)
      fs.appendFile('case_origin_owner_timestamp_csatscore.csv', rearranged_start_date_time, function(error) {
  		    if (error) throw error;
  		});
      var rearranged_end_date_time = rearranged_date + "T" + time + ":00.000Z" + ","
      console.log("call end time: " + rearranged_end_date_time)
      fs.appendFile('case_origin_owner_timestamp_csatscore.csv', rearranged_end_date_time, function(error) {
  		    if (error) throw error;
  		});

      /*
      extract csat score and store in csv file
      */
      var csat_score = current_row_data[4] + "\n"
      fs.appendFile('case_origin_owner_timestamp_csatscore.csv', csat_score, function(error) {
  		    if (error) throw error;
  		});
    }
});

/*
row_number = 1
row = data[0, row_number]
raw_date_time = row[3]
//console.log(raw_date_time)

dd_mm_yy_hh_mm_array = raw_date_time.split(" ")
dd_mm_yy = dd_mm_yy_hh_mm_array[0]
dd_mm_yy_wo_dash = dd_mm_yy.replace(/\//g, '-')
split_dd_mm_yy = dd_mm_yy_wo_dash.split("-")
correct_yyyy_mm_dd_format = split_dd_mm_yy[2] + "-" + split_dd_mm_yy[1] + "-" + split_dd_mm_yy[0]
//console.log(correct_yyyy_mm_dd_format)

hh_mm = dd_mm_yy_hh_mm_array[1]
//console.log(dd_mm_yy)

twelve_hr_format = convertTime(hh_mm, 'HH:MM')
drop_am_pm = twelve_hr_format.split(" ")
hh_mm_without_am_pm = drop_am_pm[0]

end_date_time_adjust = hh_mm_without_am_pm.split(":")
meh = end_date_time_adjust[1]
end_date_time_adjusted = end_date_time_adjust[1] - 10

newThis = hh_mm_without_am_pm.replace(meh, end_date_time_adjusted)
console.log(newThis)
hh_mm_without_am_pm = "T" + hh_mm_without_am_pm + ":00.000Z"
//console.log(hh_mm_without_am_pm)

joined = correct_yyyy_mm_dd_format + hh_mm_without_am_pm
//console.log(joined_string)
end_date_time = joined
*/
