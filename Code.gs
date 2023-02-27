// code for availability calculator and sorting method
let ss = SpreadsheetApp.getActiveSpreadsheet();

let av = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Availability")
let month_current_av = av.getRange("A2").getValue()

let av_range = av.getRange("B4:AG10")

let ovr = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Overview")
let month_current_ovr = ovr.getRange("C2").getValue()

// sets trigger to call methods automatically on a regular basis

function scriptSetup(){
  createStartupFunction();
}

function createStartupFunction(){
  ScriptApp.newTrigger("avail")
    .timeBased()
    .everyMinutes(15)
    .create()

  ScriptApp.newTrigger("sort")
    .timeBased()
    .atHour(4)
    .everyDays(1)
    .create()
}

// name of item, position in the availability table, cell of quantity in the inventory sheet
let dict = {
  "Chair" : [1, "H5"],
  "Dishes" : [2, "H12"],
  "Table 6ft" : [3, "H8"],
  "Chair Kids" : [4, "H6"],
  "Table 4ft" : [5, "H7"],
  "Table ONE" : [6, "H9"]
}


// functions to easy access and edit 2D-Array
function get_range_mon(row, col){
  return arr_mon[row-1][col-1]
}
function get_av_range(row, col){
  return arr_av[row-1][col-1]
}
function set_av_range(row, col, value){
  arr_av[row-1][col-1] = value
}
function get_id_unsorted(row, col){
  return arr_id_unsorted[row-1][col-1]
}

// counts the issued amounts of items from the booking table to display remaining amount in the availability sheet
function availablityCalc(month, item_name, av_item_index, inv_quantity) {
  let inv = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Inventory")
  let nr_columns = 41

  console.log("calculating for", item_name)

  // initialised
  quantity = inv.getRange(inv_quantity).getValue()

  for (let k=1; k<=31; k++){
    set_av_range(av_item_index, k, quantity)
  }

  // counts revenue per item
  rev_item = 0

  // loop through every columns
  for(let i=1; i <= nr_columns; i++){

    // if start & end date exist
    if (get_range_mon(1,i) !== "" && get_range_mon(5,i) !== ""){
      // make array with start date & end date
      let start_date = get_range_mon(1,i)
      let end_date = get_range_mon(5,i)
      let date_list = [start_date, end_date]

      // extracting days as int numbers into array called "days"
      let days = []
      date_list.forEach(function(date){
        day = date.toString().split(" ")[2]   // format: Sat Sep 17 2022 00:00:00 GMT-0400 (Eastern Daylight Time)
        days.push(parseInt(day))
      })

      // extending array "days" with days between start & end date
      let day_list = []
      if (days[0] === days[1]){
        day_list.push(days[0])
      } else {
        for(let j = days[0]; j <= days[1]; j++){
          day_list.push(j)
        }
      }

      // checking item slots in the booking table
      // sums the quantity of the item
      item_indices = [15, 19, 23, 27, 31, 35, 39, 43]
      nr_items = 0
      item_indices.forEach(function(index){
        if (get_range_mon(index,i) === item_name){
          nr_items += get_range_mon(index+1,i)
          rev_item += get_range_mon(index+3,i)  // sums up revenue for this item
        }
      })

      // insert in av table
      // substracts the counted amount from the previous amount
      day_list.forEach(function(day){
        let prev = get_av_range(av_item_index, day)
        set_av_range(av_item_index, day, parseInt(prev)-nr_items)
      })

    }
  }

  // sets revenue per item value in table (column "AG")
      set_av_range(av_item_index, 32, rev_item)

}

// calls the availibility calculators for the items on the dict (and some simple loading field + a "refreshed on" time stamp)
function avail(){
  let mon_av = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(month_current_av)
  var range_mon_av = mon_av.getRange("C10:AQ55")

  let dict_counter = 1
  let dict_len = Object.keys(dict).length

  // cache range_mon values into a 2D-Array
  arr_mon = range_mon_av.getValues()
  // cache av_range values into a 2D-Array
  arr_av = av_range.getValues()

  for (let key in dict){
    availablityCalc(month_current_av, key, dict[key][0], dict[key][1])
    dict_counter++
  }
  av_range.setValues(arr_av)
  av.getRange("E1").getCell(1, 1).setValue(new Date().toLocaleString("en-CA"))
}

// sorts columns on booking sheet
let number_of_columns = 41
let range_sorted_ids = "B6:B47" // on overview page

function sort(){
  let mon = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(month_current_ovr)

  let sorted_bookings = []
  let id_sorted = ovr.getRange(range_sorted_ids).getValues() // takes the ID's from sorted list on overview page to know the order

  // entering sorted IDs in array
  for (let m = 1; m<=number_of_columns; m++){
    if (id_sorted[m-1][0] !== ""){
      sorted_bookings.push(id_sorted[m-1][0])
    }
  }

  // states if any column has been moved
  let moved = false
  // counts number of sorted columns
  let index_counter = 1

  let id_unsorted = mon.getRange("C4:AQ4")
  arr_id_unsorted = id_unsorted.getValues()

  sorted_bookings.forEach(function(id){
    console.log(index_counter)
    if (moved){
      let id_unsorted = mon.getRange("C4:AQ4")
      arr_id_unsorted = id_unsorted.getValues()
    }
    for (let o = 1; o <= number_of_columns; o++){
      if (get_id_unsorted(1, o) === id && o !== index_counter){
        console.log("move column", o)
        if (!moved){
          mon.getRange("C1:I1").moveTo(ovr.getRange("D1"))
        }
        moved = true
        mon.moveColumns(id_unsorted.getCell(1, o), index_counter+2) // +2 because first 2 columns are pinned and not bookings
      }
    }
    index_counter++
  })
  if (moved){
    ovr.getRange("D1:J1").moveTo(mon.getRange("C1"))
  }

  // inserting correct formulas for statistic since they got messed up when moving columns
  if(moved){
    refreshOverview()
  }
}

function refreshOverview(){
  ovr.getRange("D2").getCell(1,1).setValue("=SUM('"+ month_current_ovr +"'!C18:AQ18)")
  ovr.getRange("F2").getCell(1,1).setValue("=SUMIF('"+ month_current_ovr +"'!C9:AQ9, \"Delivery\", '"+ month_current_ovr +"'!C18:AQ18)")
  ovr.getRange("H2").getCell(1,1).setValue("=SUMIF('"+ month_current_ovr +"'!C9:AQ9, \"Pick Up\", '"+ month_current_ovr +"'!C18:AQ18)")
  ovr.getRange("D3").getCell(1,1).setValue("=SUM('"+ month_current_ovr +"'!C20:BJ20)")
  ovr.getRange("B6").getCell(1,1).setValue("=sort(TRANSPOSE('"+ month_current_ovr +"'!C4:AQ15), TRANSPOSE('"+ month_current_ovr +"'!C10:AQ10), TRUE)")
  console.log("sorting done!")
}

function createMonthSheet(month_nr, year){
  nr_to_month = {
    '01': 'January',
    '02': 'February',
    '03': 'March',
    '04': 'April',
    '05': 'May',
    '06': 'June',
    '07': 'July',
    '08': 'August',
    '09': 'September',
    '10': 'October',
    '11': 'November',
    '12': 'December'
}

  var sheetToDuplicate = ss.getSheetByName("template");
  var monthSheet = sheetToDuplicate.copyTo(ss);
  month = nr_to_month[month_nr]
  monthSheet.setName(month + " " + year);

  var rangeToReplace = monthSheet.getRange("C2:BJ2");
  var valuesToReplace = rangeToReplace.getValues();
  for (var i = 0; i < valuesToReplace[0].length; i++) {
    valuesToReplace[0][i] = valuesToReplace[0][i].replace("MM", month_nr);
  }
  rangeToReplace.setValues(valuesToReplace);

  monthSheet.getRange("A1").setValue(month);  // Replace 'Monat' with 'April'
}

function create(){
  createMonthSheet('05', '23')  // creates a sheet for May 2023
}
