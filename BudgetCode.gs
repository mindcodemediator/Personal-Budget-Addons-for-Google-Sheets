const PAYDAYS_RANGE_NAME = "PaydaysDataRange";
function calculateBreakEvenIncome(knownIncome, expenses, hoursPerWeek, beforeTaxes) {
  // Simplified tax rate and social security rate
  var taxRate = 0.15; // Estimated 15% tax rate
  var socialSecurityRate = 0.0765; // Estimated 7.65% social security rate

  // Calculate the hourly wage needed to cover expenses before taxes
  var hourlyWageNeededBeforeTaxes = (expenses - knownIncome) / (hoursPerWeek * 2);
  if(beforeTaxes == true){
    // Calculate the hourly wage needed after taxes
    return hourlyWageNeededBeforeTaxes / (1 - taxRate - socialSecurityRate);
  } else {
    return hourlyWageNeededBeforeTaxes;
  }
}
function calculateBudgetAndPrint() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = spreadsheet.getSheetByName("Expenses"); 
  var targetSheet = spreadsheet.getSheetByName("Current Budget");

  var lastRow = sourceSheet.getLastRow();
  var sourceData = sourceSheet.getRange(1, 1, lastRow, 11).getValues(); // Retrieves data from columns A through J(1 to 9)

  var arrBills = [];
  var arrMonEss = [];
  var arrMonNonEss = [];
  var arrWeekly = [];
  var arrCredit = [];
  var arrAnnual = [];

  var arrBillInfoObjects = [];

  var paychecks = getPayDaysForMonth();
  var adjustedRow = [];

  for (var i = 1; i < sourceData.length; i++) {
    if(sourceData[i][0] != ""){
      var billInfo = {
        billName : sourceData[i][0] || "",//column A,
        category: sourceData[i][1] || "", //dropdown is in column B,
        frequency :  sourceData[i][3] || "", //column D,
        intervalAmt:  sourceData[i][4] || 0, //column E,
        amtFirstPayday : 0,
        amtSecondPayday : 0,
        totalFirstPayday : 0,
        totalSecondPayday : 0,
        weekly : 0,
        monthly : 0,
        annually : 0,
        span :  sourceData[i][6] || "", // column G
        typeSpan :  sourceData[i][7] || "", // column H
        dates: [],
        dayOfMonth: sourceData[i][2] || "", //column C
        lastPaid: sourceData[i][5] || "", //columnF
        dueFirstPayPeriod: "",
        dueSecondPayPeriod: "",
        dayOfWeek: sourceData[i][9] || "", //column J
        apptRange:  sourceData[i][8] || "", //column I
        saveOtherPaycheck: sourceData[i][10] || 0  //column K
      }
      if(billInfo.billName != ""){
        switch(billInfo.frequency){
          case "Weekly":
            billInfo = setWeeklyInfo(billInfo, paychecks);
            break;
          case "Daily":
            billInfo = setDailyInfo(billInfo, paychecks);
            break;
          case "Monthly":
            billInfo = setMonthlyInfo(billInfo, paychecks);
            break;
          case "Annually":
            billInfo.span = 12;
            billInfo.typeSpan = "Months";
            billInfo = setMonthlyInterval(billInfo, paychecks);
            break;
          case "Interval":
            if(billInfo.typeSpan == "Days" || billInfo.typeSpan == "Weeks"){
              billInfo = setWeeklyAndDailyInterval(billInfo, paychecks);
            } else {
              billInfo = setMonthlyInterval(billInfo, paychecks);
            }
            break;
          default:
            Logger.log("ERROR with frrequency type");
            break;
        }
        arrBillInfoObjects.push(billInfo);
        adjustedRow = [billInfo.billName, billInfo.dates, billInfo.frequency, billInfo.amtFirstPayday,billInfo.amtSecondPayday, billInfo.weekly, billInfo.monthly, billInfo.annually, billInfo.span, billInfo.typeSpan, billInfo.totalFirstPayday, billInfo.totalSecondPayday, billInfo.intervalAmt];
        switch(billInfo.category){
          case "Bills":
            arrBills.push(adjustedRow);
            break;
          case "Monthly Essentials":
            arrMonEss.push(adjustedRow);
            break;
          case "Monthly Non-Essentials":
            arrMonNonEss.push(adjustedRow);
            break;
          case "Weekly":
            arrWeekly.push(adjustedRow);
            break;
          case "Credit":
            arrCredit.push(adjustedRow);
            break;
          case "Annual":
            arrAnnual.push(adjustedRow);
            break;
          default:
            Logger.log("ERROR with bill category");
            break;
        }
      }
    }    
  }

  // Now you have three separate arrays based on dropdown values
  // arrOption1, arrOption2, and arrOption3

  // You can choose to do whatever you want with these arrays
  // For example, you can set the values in different named ranges
  setValuesToNamedRange(spreadsheet, "ChartBillsRange", arrBills);
  setValuesToNamedRange(spreadsheet, "ChartMonthlyEssentialsRange", arrMonEss);
  setValuesToNamedRange(spreadsheet, "ChartMonthlyNonEssentialsRange", arrMonNonEss);
  setValuesToNamedRange(spreadsheet, "ChartWeeklyExpensesRange", arrWeekly);
  setValuesToNamedRange(spreadsheet, "ChartCreditRange", arrCredit);
  setValuesToNamedRange(spreadsheet, "ChartAnnualRange", arrAnnual);


  printMonthlyDetails(arrBillInfoObjects, paychecks);
   
  // Lock the sheet as read-only
  var protection = targetSheet.protect().setDescription('Read Only Protection');
  
  protection.setWarningOnly(true); // Display a warning message when someone tries to edit

 // Wait for a few seconds (if needed) before setting values
 // Utilities.sleep(5000); // Adjust the sleep time as needed
}

function printMonthlyDetails(billInfoArray, paychecks){

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("Current Budget");
  
  var paycheckOneRange = spreadsheet.getRangeByName("PaycheckOneRange");
  var paycheckTwoRange = spreadsheet.getRangeByName("PaycheckTwoRange");

  var billAcctRange = spreadsheet.getRangeByName("BillAcctRange");
  var spendAcctRange = spreadsheet.getRangeByName("SpendAcctRange");
  var savingsAcctRange = spreadsheet.getRangeByName("SavingsAcctRange");
  var neededInAcctRange = spreadsheet.getRangeByName("NeededInAcctRange");

  // Clear the contents of the named range
  paycheckOneRange.clearContent();
  paycheckTwoRange.clearContent();

  
  // Calculate the dynamic starting cells for the first and second paycheck columns
  var startingCellForFirstPaycheck = calculateStartingCell(sheet, "First Paycheck");
  var startingCellForSecondPaycheck = calculateStartingCell(sheet, "Second Paycheck");

  //VALERIE WORKING HERE
  var startingCellForFirstTotal = calculateStartingCellRight(sheet, "First Total");
  var startingCellForSecondTotal = calculateStartingCellRight(sheet, "Second Total");
  var startingCellForFirstPayDate = calculateStartingCellRight(sheet, "First Paycheck");
  var startingCellForSecondPayDate = calculateStartingCellRight(sheet, "Second Paycheck");
  var startingCellTotalNeeded = calculateStartingCell(sheet, "Total Needed");
  var extraMoneyA1Notation = calculateStartingCell(sheet, "Amount extra");
  var cell =sheet.getRange(extraMoneyA1Notation);
  var extraMoney = cell.getValue();
  //var cell = sheet.getRange(incomeTotalA1Notation);
  //var incomeTotal = cell.getValue();
  

  sheet.getRange(startingCellForFirstTotal).clearContent();
  sheet.getRange(startingCellForSecondTotal).clearContent();
  sheet.getRange(startingCellForFirstPayDate).clearContent();
  sheet.getRange(startingCellForSecondPayDate).clearContent();


  var firstPaycheckBills = [];
  var secondPaycheckBills = [];
  var printArray = [];
  var billAcctFirst = 0;
  var billAcctSecond = 0;
  var spendAcctFirst = 0;
  var spendAcctSecond = 0;
  var billAcctTotalFirst = 0;
  var billAcctTotalSecond = 0;

  billInfoArray.forEach(function(billInfo, index){
    if(billInfo.amtFirstPayday > 0){
      firstPaycheckBills.push([billInfo.billName, billInfo.amtFirstPayday, billInfo.dueFirstPayPeriod]);
      printArray.push([billInfo.billName, billInfo.dates[0],  billInfo.frequency, billInfo.amtFirstPayday, billInfo.dueFirstPayPeriod, 1]);
      for(var i = 1; i < billInfo.dates.length; i++){
        if(billInfo.dates[i] >= paychecks[0] && billInfo.dates[i] < paychecks[1]){
          if(billInfo.saveOtherPaycheck != 0 && billInfo.dueFirstPayPeriod != ""){
            printArray.push([billInfo.billName, billInfo.dates[i],  billInfo.frequency, billInfo.amtFirstPayday + billInfo.saveOtherPaycheck, billInfo.dueFirstPayPeriod, 1]);  
          } else {
            printArray.push([billInfo.billName, billInfo.dates[i],  billInfo.frequency, billInfo.amtFirstPayday, billInfo.dueFirstPayPeriod, 1]);
          }
        }
      }
    }
    if(billInfo.amtSecondPayday > 0){
      secondPaycheckBills.push([billInfo.billName, billInfo.amtSecondPayday, billInfo.dueSecondPayPeriod]);
      printArray.push([billInfo.billName, billInfo.dates[0],  billInfo.frequency, billInfo.amtSecondPayday, billInfo.dueSecondPayPeriod, 2]);
      for(var i = 1; i < billInfo.dates.length; i++){
        if(billInfo.dates[i] >= paychecks[1] && billInfo.dates[i] < paychecks[2]){
          if(billInfo.saveOtherPaycheck != 0 && billInfo.dueFirstPayPeriod != ""){
            printArray.push([billInfo.billName, billInfo.dates[i],  billInfo.frequency, billInfo.amtSecondPayday + billInfo.saveOtherPaycheck, billInfo.dueSecondPayPeriod, 2]);
          } else {
            printArray.push([billInfo.billName, billInfo.dates[i],  billInfo.frequency, billInfo.amtSecondPayday, billInfo.dueSecondPayPeriod, 2]);
          }
        }
      }
    }
    if(billInfo.frequency == "Weekly" || billInfo.frequency ==  "Daily"){
      spendAcctFirst = spendAcctFirst + billInfo.amtFirstPayday;
      spendAcctSecond = spendAcctSecond + billInfo.amtSecondPayday;
    } else {
      billAcctFirst = billAcctFirst + billInfo.amtFirstPayday;
      billAcctSecond = billAcctSecond + billInfo.amtSecondPayday;
      billAcctTotalFirst = billAcctTotalFirst + billInfo.totalFirstPayday;
      billAcctTotalSecond = billAcctTotalSecond + billInfo.totalSecondPayday;
    }
  });
  var totalFirst = billAcctFirst + spendAcctFirst + extraMoney;
  var totalSecond = billAcctSecond + spendAcctSecond + extraMoney;
  var biggerAmt = 0;
  var savingsPayPeriod = 0;
  var savingsAmt = [];
  if(totalFirst > totalSecond){
    biggerAmt = totalFirst;
    savingsPayPeriod = 2;
    savingsAmt = [0,totalFirst - totalSecond];
  } else {
    biggerAmt = totalSecond;
    savingsPayPeriod = 1;
    savingsAmt = [totalSecond - totalFirst, 0];
  }

  updateRangeDynamically(sheet, startingCellForFirstPaycheck, firstPaycheckBills);
  updateRangeDynamically(sheet, startingCellForSecondPaycheck, secondPaycheckBills);
  sheet.getRange(startingCellForFirstTotal).setValue(totalFirst - extraMoney);
  sheet.getRange(startingCellForSecondTotal).setValue(totalSecond - extraMoney);
  sheet.getRange(startingCellForFirstPayDate).setValue(paychecks[0]);
  sheet.getRange(startingCellForSecondPayDate).setValue(paychecks[1]);
  sheet.getRange(startingCellTotalNeeded).setValue(biggerAmt);
  printArrayHorizontally(billAcctRange, [billAcctFirst,billAcctSecond]);
  printArrayHorizontally(spendAcctRange, [spendAcctFirst, spendAcctSecond]);
  printArrayHorizontally(savingsAcctRange, savingsAmt);
  printArrayHorizontally(neededInAcctRange, [billAcctTotalFirst, billAcctTotalSecond]);


  //Then copy everything to separate sheet for app to use
  var printSheet = spreadsheet.getSheetByName("AppTotals");
  var startingCellPrint = calculateStartingCell(printSheet, "Bill Name");
  clearCells(printSheet, startingCellPrint);
  updateRangeDynamically(printSheet, startingCellPrint, printArray);
  copyValuesAndFormatting();
}



function setDailyInfo(dailyBillInfo, paydays){
    var daysBetweenFirst = daysBetweenDates(paydays[0], paydays[1]);
    var daysBetweenSecond = daysBetweenDates(paydays[1], paydays[2]);
    dailyBillInfo.amtFirstPayday = daysBetweenFirst * dailyBillInfo.intervalAmt;
    dailyBillInfo.amtSecondPayday = daysBetweenSecond* dailyBillInfo.intervalAmt;
    dailyBillInfo.totalFirstPayday = dailyBillInfo.amtFirstPayday;
    dailyBillInfo.totalSecondPayday = dailyBillInfo.amtSecondPayday;
    dailyBillInfo.weekly = dailyBillInfo.intervalAmt * 7;
    dailyBillInfo.monthly = dailyBillInfo.intervalAmt * 30.5;
    dailyBillInfo.annually = dailyBillInfo.intervalAmt * 365;
    dailyBillInfo.dueFirstPayPeriod = "Due";
    dailyBillInfo.dueSecondPayPeriod = "Due";
    dailyBillInfo.dates.push("");

    return dailyBillInfo;
}

function setWeeklyInfo(weeklyBillInfo, paydays){
  var countFirst = countDaysOfWeekBetweenDates(paydays[0], paydays[1], weeklyBillInfo.dayOfWeek);
  var countSecond = countDaysOfWeekBetweenDates(paydays[1], paydays[2], weeklyBillInfo.dayOfWeek);
  weeklyBillInfo.amtFirstPayday = countFirst * weeklyBillInfo.intervalAmt;
  weeklyBillInfo.amtSecondPayday = countSecond* weeklyBillInfo.intervalAmt;
  weeklyBillInfo.totalFirstPayday = weeklyBillInfo.amtFirstPayday;
  weeklyBillInfo.totalSecondPayday = weeklyBillInfo.amtSecondPayday;
  weeklyBillInfo.weekly = weeklyBillInfo.intervalAmt;
  weeklyBillInfo.monthly = (weeklyBillInfo.intervalAmt / 7) * 30.5;
  weeklyBillInfo.annually = (weeklyBillInfo.intervalAmt / 7) * 365;
  weeklyBillInfo.dueFirstPayPeriod = "Due";
  weeklyBillInfo.dueSecondPayPeriod = "Due";
  weeklyBillInfo.dates.push("");

  return weeklyBillInfo;
}

function setMonthlyInfo(monthlyBillInfo, paydays){ 
  if(monthlyBillInfo.apptRange != ""){
    monthlyBillInfo.intervalAmt = getAppointments(monthlyBillInfo.dayOfMonth ,monthlyBillInfo.intervalAmt, monthlyBillInfo.apptRange);
  }
  var dayOfMonthFirst = paydays[0].getDate();
  var dayOfMonthSecond = paydays[1].getDate();
  var now = new Date();
  var amtDuePaycheck = monthlyBillInfo.intervalAmt;

  if(monthlyBillInfo.dayOfMonth < dayOfMonthFirst){
    //second pay period but day is less than first paycheck so this actually gets paid for the next month
    monthlyBillInfo.amtFirstPayday = monthlyBillInfo.saveOtherPaycheck;
    monthlyBillInfo.amtSecondPayday = monthlyBillInfo.intervalAmt - monthlyBillInfo.saveOtherPaycheck;
    monthlyBillInfo.totalFirstPayday = monthlyBillInfo.saveOtherPaycheck;
    monthlyBillInfo.totalSecondPayday = monthlyBillInfo.intervalAmt;
    monthlyBillInfo.dueSecondPayPeriod = "Due";
    var setDate = new Date(now.getFullYear(), now.getMonth() + 1, monthlyBillInfo.dayOfMonth);
    monthlyBillInfo.dates.push(setDate);
  } else if(monthlyBillInfo.dayOfMonth > dayOfMonthSecond){
    //second pay period
    monthlyBillInfo.amtFirstPayday = monthlyBillInfo.saveOtherPaycheck;
    monthlyBillInfo.amtSecondPayday = monthlyBillInfo.intervalAmt - monthlyBillInfo.saveOtherPaycheck;
    monthlyBillInfo.totalFirstPayday = monthlyBillInfo.saveOtherPaycheck;
    monthlyBillInfo.totalSecondPayday = monthlyBillInfo.intervalAmt;
    monthlyBillInfo.dueSecondPayPeriod = "Due";
    var setDate = new Date(now.getFullYear(), now.getMonth(), monthlyBillInfo.dayOfMonth);
    monthlyBillInfo.dates.push(setDate);
  } else {
    //first pay period
    monthlyBillInfo.amtFirstPayday = monthlyBillInfo.intervalAmt - monthlyBillInfo.saveOtherPaycheck;
    monthlyBillInfo.amtSecondPayday = monthlyBillInfo.saveOtherPaycheck;
    monthlyBillInfo.totalFirstPayday = monthlyBillInfo.intervalAmt;
    monthlyBillInfo.totalSecondPayday = monthlyBillInfo.saveOtherPaycheck;
    monthlyBillInfo.dueFirstPayPeriod = "Due";
    var setDate = new Date(now.getFullYear(), now.getMonth(), monthlyBillInfo.dayOfMonth);
    monthlyBillInfo.dates.push(setDate);
  }
  monthlyBillInfo.weekly = monthlyBillInfo.intervalAmt / getWeeksBetween(paydays[0], paydays[2]);
  monthlyBillInfo.monthly = monthlyBillInfo.intervalAmt;
  monthlyBillInfo.annually = monthlyBillInfo.intervalAmt * 12;
  return monthlyBillInfo;
}

function setMonthlyInterval(monthlyIntervalBillInfo, paychecks){
   var span = monthlyIntervalBillInfo.span;
  
  var dueDate = new Date(monthlyIntervalBillInfo.lastPaid);
  var prevDate = new Date(dueDate);
  prevDate = addOrSubMonths(prevDate, -1 * span)
  var now = new Date();
  now.setDate(dueDate.getDate());

  var paydaysInFirstInterval = [];
  var paydaysInSecondInterval = [];

  if(now < paychecks[1]){
    //to make sure we are checking from within the pay period. if it is less than the first paycheck.
    now = addOrSubMonths(now,1);
  }
  if(dueDate >= paychecks[2]){
    //if the date entered in the spreadsheet is later, check to see if it is within interval
    //if not,  then nothing is owed

    var tempDate = addOrSubMonths(now, span);
    if(tempDate < dueDate){
      //its not due in this interval pay nothing right now, set aside nothing right now
      monthlyIntervalBillInfo.amtFirstPayday = 0;
      monthlyIntervalBillInfo.amtSecondPayday = 0;
      monthlyIntervalBillInfo.totalFirstPayday = 0;
      monthlyIntervalBillInfo.totalSecondPayday = 0;
      var numWeeks = getWeeksBetween(now, tempDate);
      monthlyIntervalBillInfo.weekly = monthlyIntervalBillInfo.intervalAmt / numWeeks;
      monthlyIntervalBillInfo.monthly = monthlyIntervalBillInfo.intervalAmt / span;
      monthlyIntervalBillInfo.annually = (monthlyIntervalBillInfo.intervalAmt / span) * 12;
      monthlyIntervalBillInfo.dates.push(dueDate);
      return monthlyIntervalBillInfo;
    }
  } else if(dueDate < paychecks[0]){
    //otherwise, if the due date is before current pay period, find the next due date
    prevDate = new Date(dueDate);
    dueDate = addOrSubMonths(dueDate, span);
    while(dueDate < paychecks[0]){
      prevDate = new Date(dueDate);
      dueDate = addOrSubMonths(dueDate, span);
    }
  }
  var numWeeks = getWeeksBetween(prevDate, dueDate);
  monthlyIntervalBillInfo.weekly = monthlyIntervalBillInfo.intervalAmt / numWeeks;
  monthlyIntervalBillInfo.monthly = monthlyIntervalBillInfo.intervalAmt / span;
  monthlyIntervalBillInfo.annually = (monthlyIntervalBillInfo.intervalAmt / span) * 12;
  monthlyIntervalBillInfo.dates.push(dueDate);
  if(dueDate < paychecks[1]){
    monthlyIntervalBillInfo.dueFirstPayPeriod = "Due";
  } else if(dueDate < paychecks[2]){
    monthlyIntervalBillInfo.dueSecondPayPeriod = "Due";
  }

//Get amount for first payday
  paydaysInFirstInterval = getPayDaysInInterval(prevDate, dueDate);
  monthlyIntervalBillInfo.amtFirstPayday = monthlyIntervalBillInfo.intervalAmt / paydaysInFirstInterval.length;
  var totalFirst = 0;
  var i = 0;
  while(paydaysInFirstInterval[i] <= paychecks[0]){
    totalFirst = totalFirst + monthlyIntervalBillInfo.amtFirstPayday;
    i = i + 1;
  }
  
  monthlyIntervalBillInfo.totalFirstPayday = totalFirst;

  if(dueDate >= paychecks[1]){
    //not in first pay period, so the amount is the same for both pay days
    monthlyIntervalBillInfo.amtSecondPayday = monthlyIntervalBillInfo.amtFirstPayday;
    monthlyIntervalBillInfo.totalSecondPayday = monthlyIntervalBillInfo.totalFirstPayday + monthlyIntervalBillInfo.amtFirstPayday;
    if(dueDate < paychecks[2]){
      //if it is due, show full amount
      monthlyIntervalBillInfo.amtSecondPayday = monthlyIntervalBillInfo.totalSecondPayday;
    }
  } else {
    //due in first pay period, get interval and information for second payday
    //total will be the same because this will be the first payment in the interval
    monthlyIntervalBillInfo.amtFirstPayday = totalFirst;
    prevDate = new Date(dueDate);
    dueDate = addOrSubMonths(dueDate, span);
    paydaysInSecondInterval = getPayDaysInInterval(prevDate, dueDate);
    monthlyIntervalBillInfo.amtSecondPayday = monthlyIntervalBillInfo.intervalAmt / paydaysInSecondInterval.length;
    monthlyIntervalBillInfo.totalSecondPayday = monthlyIntervalBillInfo.amtSecondPayday;
  }
  
  return monthlyIntervalBillInfo; 

}

function setWeeklyAndDailyInterval(weeklyIntervalBillInfo, paychecks){
  var span = weeklyIntervalBillInfo.span;

  if(weeklyIntervalBillInfo.typeSpan == "Weeks" + weeklyIntervalBillInfo.span == 1){
    //set day of week first
    weeklyIntervalBillInfo.dayOfWeek = getDayOfWeek(weeklyIntervalBillInfo.lastPaid);
    //if it is every 1 week, just do a weekly function
    return setWeeklyInfo(weeklyIntervalBillInfo);
  }

  
  var addDays;
  if(weeklyIntervalBillInfo.typeSpan == "Days"){
    //daily
    addDays = weeklyIntervalBillInfo.span;
  } else {
    //weekly
    addDays = weeklyIntervalBillInfo.span * 7;
  }
  var costPerDay = weeklyIntervalBillInfo.intervalAmt / addDays;
    
  weeklyIntervalBillInfo.weekly = costPerDay * 7;
  weeklyIntervalBillInfo.annually = weeklyIntervalBillInfo.weekly * 52;
  weeklyIntervalBillInfo.monthly = weeklyIntervalBillInfo.annually / 12;

  var dueDate = new Date(weeklyIntervalBillInfo.lastPaid);
  var prevDate = new Date(dueDate);
  prevDate.setDate(prevDate.getDate() - addDays);
  var now = new Date();
  //check to see if due date is in the future
  if(now < paychecks[1]){
    //to make sure we are checking from within the pay period. if it is less than the first paycheck.
    now = addOrSubMonths(now, 1);
  }
  if(dueDate > paychecks[2]){
    //if the date entered in the spreadsheet is later, check to see if it is within interval
    //if not,  then nothing is owed
    var tempDate = new Date(now);
    tempDate.setDate(tempDate.getDate() + addDays);
    if(tempDate < dueDate){
      //its not due in this interval pay nothing right now, set aside nothing right now
      weeklyIntervalBillInfo.amtFirstPayday = 0;
      weeklyIntervalBillInfo.amtSecondPayday = 0;
      weeklyIntervalBillInfo.totalFirstPayday = 0;
      weeklyIntervalBillInfo.totalSecondPayday = 0;
      weeklyIntervalBillInfo.dates.push(dueDate);
      return weeklyIntervalBillInfo;
    }
  } else if(dueDate < paychecks[0]){
    //if it is before the pay period, advance it.
    prevDate = new Date(dueDate);
    dueDate.setDate(dueDate.getDate() + addDays);
    while(dueDate < paychecks[0]){
      //loop over dates until we find a due date for after the first paycheck of the month
      prevDate = new Date(dueDate);
      dueDate.setDate(dueDate.getDate() + addDays);
    }
  }  
    //Get amount for first payday
  var amtPerPayDay = [];
  var totalAmtPerPayDay = [];
  var dueInPeriod = [];
  amtPerPayDay.push(0,0);
  totalAmtPerPayDay.push(0,0);
  dueInPeriod.push("", "");

  //we are calculating cost per day so it doesn't matter 
  //how many fall in between each paycheck
  
  var daysSincePrev = [];
  var daysInPayPeriod = [];

  for(var i = 0; i < paychecks.length - 1; i++){
    daysSincePrev.push(daysBetweenDates(prevDate, paychecks[i]));
    daysInPayPeriod.push(daysBetweenDates(paychecks[i], paychecks[i+1]));
    amtPerPayDay[i] = daysInPayPeriod[i] * costPerDay;
    totalAmtPerPayDay[i] = (daysSincePrev[i] *  costPerDay) + amtPerPayDay[i];
    if(dueDate< paychecks[i+ 1]){
      //if the due date falls between paychecks, add the date to the array and move the prev date
      weeklyIntervalBillInfo.dates.push(dueDate);
      dueInPeriod[i] = "Due";
      //advance Due date and prev date
     // prevDate = new Date(dueDate); 
      dueDate.setDate(dueDate.getDate() + addDays);
      while(dueDate < paychecks[i + 1]){
        prevDate = new Date(dueDate); 
        dueDate.setDate(dueDate.getDate() + addDays);
      }
    }
  }
  //if the date waasn't due this month, push the due date
  if(weeklyIntervalBillInfo.dates.length == 0){
    weeklyIntervalBillInfo.dates.push(dueDate);
  }

  //fill out the rest of the info in the object
  weeklyIntervalBillInfo.amtFirstPayday = amtPerPayDay[0];
  weeklyIntervalBillInfo.amtSecondPayday = amtPerPayDay[1];
  weeklyIntervalBillInfo.totalFirstPayday = totalAmtPerPayDay[0];
  weeklyIntervalBillInfo.totalSecondPayday = totalAmtPerPayDay[1];
  weeklyIntervalBillInfo.dueFirstPayPeriod = dueInPeriod[0];
  weeklyIntervalBillInfo.dueSecondPayPeriod = dueInPeriod[1];
  return weeklyIntervalBillInfo;
}


function getAppointments(dayOfBill, costPerSession, rangeName){
  //get last date paid
  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var AppointmentsRange = spreadsheet.getRangeByName(rangeName);
  var apptsData = AppointmentsRange.getValues();
  var compareDate;
  var total = 0;
  

  var before = apptsData[0][0] == "Before" ? true : false;
  var lastTimeDue;
  var thisTimeDue;
  if(before){
    lastTimeDue = createDateForMonth(dayOfBill, -1);
    thisTimeDue = createDateForMonth(dayOfBill, 0);
  } else{
    lastTimeDue = createDateForMonth(dayOfBill, 0);
    thisTimeDue = createDateForMonth(dayOfBill, 1);
  }
  for(var i = 1; i < apptsData.length; i++){
    if(apptsData[i][0] != ""){
      compareDate = new Date(apptsData[i][0]);
      if(compareDate >= lastTimeDue && compareDate < thisTimeDue){
        total = total + costPerSession;
      }
    } else {
      break;
    }
  }
  return total;
}
/////////////////////////////////////////////////////////
//HELPER FUNCTIONS FOR PRINTING
////////////////////////////////////////////////////////////
function clearCells(sheet, startCellA1Notation) {
  // Find the last filled row and column
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  
  // Get row and column numbers for the starting cell in A1 notation
  var startCell = sheet.getRange(startCellA1Notation);
  var startRow = startCell.getRow();
  var startCol = startCell.getColumn();
  
  // Calculate the number of rows and columns to clear
  var numRowsToClear = lastRow - startRow + 1;
  var numColsToClear = lastCol - startCol + 1;
  
  // Clear the range starting from the provided cell and extending to the last filled cell
  if (numRowsToClear > 0 && numColsToClear > 0) {
    sheet.getRange(startRow, startCol, numRowsToClear, numColsToClear).clear();
  }
}

function copyValuesAndFormatting() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  // Get the source and target sheets by name
  var sourceSheet = spreadsheet.getSheetByName("Current Budget");

  var currentDate = new Date();
  var monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  var currentMonth = monthNames[currentDate.getMonth()];
  var newSheetName = currentMonth + "_" + currentDate.getFullYear(); // Replace with your desired sheet name


  var targetSheet = spreadsheet.getSheetByName(newSheetName);
  if ( targetSheet == null) {
      // Create a new sheet with the specified name
      var targetSheet = spreadsheet.insertSheet(newSheetName);
  } else {
    // If it does exist, clear all contents and formatting
    targetSheet.clear({contentsOnly: true});
    targetSheet.clear({formatOnly: true});
  }

  // Check if both sheets exist
  if (sourceSheet && targetSheet) {
    // Clear all conditional formatting and data validations in the target sheet
    //targetSheet.clear({validationsOnly: true});

    // Get the data range of the source sheet (values and formatting)
    var sourceDataRange = sourceSheet.getDataRange();

    // Get the dimensions of the data range
    var numRows = sourceDataRange.getNumRows();
    var numCols = sourceDataRange.getNumColumns();

    // Get the target range in the target sheet with the same dimensions
    var targetRange = targetSheet.getRange(1, 1, numRows, numCols);

    // Copy both values and formatting from the source range to the target range
    sourceDataRange.copyTo(targetRange);

    Logger.log("Values and formatting copied to new sheet");
  } else {
    Logger.log("One or both sheets not found.");
  }
}

function calculateStartingCell(sheet, label) {
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var data = sheet.getRange(1, 1, lastRow, lastColumn).getValues();

  for (var i = 0; i < lastRow; i++) {
    for (var j = 0; j < lastColumn; j++) {
      if (data[i][j] === label) {
        return sheet.getRange(i + 2, j + 1).getA1Notation();
      }
    }
  }

  // If the label is not found, start in the next available row in Column B
  var nextRow = lastRow + 1;
  var nextColumn = 2; // Column B
  return sheet.getRange(nextRow, nextColumn).getA1Notation();
}

function calculateStartingCellRight(sheet, label) {
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var data = sheet.getRange(1, 1, lastRow, lastColumn).getValues();

  for (var i = 0; i < lastRow; i++) {
    for (var j = 0; j < lastColumn; j++) {
      if (data[i][j] === label) {
        return sheet.getRange(i + 1, j + 2).getA1Notation();
      }
    }
  }

  // If the label is not found, start in the next available row in Column B
  var nextRow = lastRow + 1;
  var nextColumn = 2; // Column B
  return sheet.getRange(nextRow, nextColumn).getA1Notation();
}
function updateRangeDynamically(sheet, startingCell, data) {
  var numRows = data.length;
  var numCols = data[0].length; // Update to consider the number of columns in the data

  var range = sheet.getRange(startingCell).offset(0, 0, numRows, numCols);
  range.clearContent(); // Clear previous content
  if (numRows > 0 && numCols > 0) {
    range.setValues(data);
  }
}
function printArrayHorizontally(namedRange, itemArray) {
  //Sizes NEED to match of range and array


  // Check if the named range exists and if it has the same number of cells as the array
  if (namedRange && namedRange.getNumRows() === 1 && namedRange.getNumColumns() === itemArray.length) {
    // Create a 2D array with a single row for the horizontal print
    var horizontalArray = [itemArray];

    // Set the values of the named range using the 2D array
    namedRange.setValues(horizontalArray);
  } else {
    Logger.log("Named range does not exist or has a different size than the array.");
  }
}
///////////////////////////////////////////////////////////////////////////////
//HELPER FUNCTIONSn  FOR CALCULATING NUMBERS//
/////////////////////////////////////////////////////////////////////////////
function setValuesToNamedRange(spreadsheet, namedRange, values) {
  var range = spreadsheet.getRangeByName(namedRange);
  if (range && values.length > 0) {
    range.clearContent()
;    // Resize the range to match the size of the values array
    range = range.offset(0, 0, values.length, values[0].length);
    try {
      range.setValues(values);
    } catch(e) {
      Logger.log("Error setting value: " + e);
    }
  }
}

function createDateForMonth(dayOfMonth, increment) {
  // Get the current date
  var currentDate = new Date();
  
  // Get the year and month of the current date
  var year = currentDate.getFullYear();
  var month = currentDate.getMonth();
  
  month = month + increment //if increment is negative, it will get the prior months
  if(month === 0){
    year -= 1;
    month = 11 //december
  } else if(month == 12){
    year = year + 1;
    month = 0;
  }
  
  
  // Create a Date object for the 10th day of the previous month
  return new Date(year, month, dayOfMonth);
}

function daysBetweenDates(start, end) {
  // Parse the date strings into Date objects
  var startDate = new Date(start);
  var endDate = new Date(end);

  // Calculate the time difference in milliseconds
  var timeDifference = endDate.getTime() - startDate.getTime();

  // Calculate the number of days
  var daysDifference = Math.floor(timeDifference / (1000 * 60 * 60 * 24));

  return daysDifference;
}


function getDayOfWeek(checkDate) {
  // Create a Date object from the given date string
  var date = new Date(checkDate);
  
  // Use getDay() to get the day of the week as a number (0-6)
  var dayIndex = date.getDay();
  
  // Array mapping day indices to names
  var days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  
  // Get the name of the day
  var dayName = days[dayIndex];
  
  // Log or return the result
  return dayName;
}

function getAmountForInterval(total, numPaychecks){
  return total/numPaychecks;
}
function getPayDaysInInterval(startDate, endDate){
  var allPaydays = getPayDays();
  var paydaysIncluded = [];
  var checkDate;
  for(var i = 0; i < allPaydays.length; i++){
    checkDate = new Date(allPaydays[i][0]);
    if(checkDate <= endDate && checkDate > startDate){
      paydaysIncluded.push(checkDate);
    }
  }
  return paydaysIncluded;
}

function getPayDays(){
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
 // Specify the named range for paydays data
  var paydaysDataRange = spreadsheet.getRangeByName(PAYDAYS_RANGE_NAME);
  var dateArray = paydaysDataRange.getValues();  
  var sortedDates = dateArray.sort((a, b) => new Date(a) - new Date(b));
  return sortedDates;
}

function getPayDaysForMonth() {

  var now = new Date();
  var currentMonth = now.getMonth();
  var currentPaydays = [];
  var found = false;
  var paydays = getPayDays();

  for (var i = 0; i < paydays.length; i++) {
    //this loop will add the two paydays for the month plus one extra after the second payday
    var payday = new Date(paydays[i][0])
    if (payday.getMonth() === currentMonth) {
      if (found) {
        //add one more to get the following payday
        currentPaydays.push(paydays[i+1][0]);
      } 
      currentPaydays.push(payday);
      found = true;
    }
  }
  var sortedPaydays = currentPaydays.sort((a, b) => new Date(a) - new Date(b));
  return sortedPaydays;
}

function daysBetweenDates(start, end) {
  // Parse the date strings into Date objects
  var startDate = new Date(start);
  var endDate = new Date(end);

  // Calculate the time difference in milliseconds
  var timeDifference = endDate.getTime() - startDate.getTime();

  // Calculate the number of days
  var daysDifference = Math.floor(timeDifference / (1000 * 60 * 60 * 24));

  return daysDifference;
}

function countDaysOfWeekBetweenDates(date1, date2, dayName) {
  // Parse the input dates to JavaScript Date objects.
  var startDate = new Date(date1);
  var endDate = new Date(date2);
  
  // Define an array of weekday names to compare against the input.
  var days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  
  // Find the index of the target day.
  var targetDayIndex = days.indexOf(dayName);
  if (targetDayIndex === -1) {
    return 'Invalid day name';
  }
  
  var count = 0;
  
  // Start iterating from the start date.
  for (var d = startDate; d < endDate; d.setDate(d.getDate() + 1)) {
    if (d.getDay() === targetDayIndex) {
      count++;
    }
  }
  
  return count;
}

function getWeeksBetween(startDate, endDate) {
  if (!startDate || !endDate) {
    return 'Please provide valid dates';
  }
  
  // Calculate the time difference in milliseconds
  var timeDifference = endDate - startDate;
  
  // Convert milliseconds to weeks with decimals (1 week = 7 days, 1 day = 24 hours * 60 minutes * 60 seconds * 1000 milliseconds)
  var weeksDifference = timeDifference / (7 * 24 * 60 * 60 * 1000);
  
  return weeksDifference;
}

function addOrSubMonths(inputDate, numMonths) {
  // Convert the input date to a JavaScript Date object
  var date = new Date(inputDate);

  // Calculate the new month and year after adding or subtracting months
  var newMonth = date.getMonth() + numMonths;
  var newYear = date.getFullYear();

  // Handle year transitions while accounting for positive and negative results
  while (newMonth < 0) {
    newYear -= 1; // Subtract a year
    newMonth += 12; // Add 12 months to wrap around to the previous year
  }
  while (newMonth > 11) {
    newYear += 1; // Add a year
    newMonth -= 12; // Subtract 12 months to wrap around to the next year
  }

  // Set the new month and year in the date object
  date.setMonth(newMonth);
  date.setFullYear(newYear);
  
  //Handle when the day of the month is after the actual number of days in the month
  var numDaysInMonth = getDaysInMonth(newYear, newMonth);
  while(date.getDate() > numDaysInMonth){
    date.setDate(date.getDate() - 1);
  }
  return date; // Return the JavaScript Date object
}

function getDaysInMonth(year, month) {
  // Create a new Date object for the given year and month (months are 0-based)
  var date = new Date(year, month, 1);

  // Move to the next month and subtract one day to get the last day of the current month
  date.setMonth(date.getMonth() + 1);
  date.setDate(date.getDate() - 1);

  // The date object now contains the last day of the month, so you can extract the day part
  return date.getDate();
}