// copyright (c) Henrik Bechmann, Toronto, 2015, all rights reserved.
// This source code may be used by anyone according to the terms of the 
//   GPL licence at http://www.gnu.org/copyleft/gpl.html
function GenerateBudgetSheets() {
  GenerateActualMatrix();
  GenerateCharts();
  GenerateChangeMatrix();
}


function GenerateActualMatrix() {
  var theSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  _setBaseMatrix(theSpreadsheet);
  _setBaseAggregates(theSpreadsheet);
}


function GenerateCharts() {
  var theSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  _setActualRefCharts(theSpreadsheet);
  _callActualTrendCharts(theSpreadsheet);
  _callConstantTrendCharts(theSpreadsheet);
  _callCommonTrendCharts(theSpreadsheet);
}


function GenerateChangeMatrix() {
  var theSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  /* ------------------------------------------
   * copy data to 'Period Changes' sheet
   * ------------------------------------------ */
  var actualsheet, changesheet, parametersheet;
  actualsheet = theSpreadsheet.getSheetByName('Consolidated Actual');
  changesheet = theSpreadsheet.getSheetByName('Period Changes');
  parametersheet = theSpreadsheet.getSheetByName('Parameters');
  var accountrowindexes, startyear, endyear, titleindexes;
  var datastartindex, dataendindex, startyearcolindex, endyearcolindex;
  accountrowindexes = actualsheet.getRange('A:A').getValues();
  accountrowindexes = _flattenValues(accountrowindexes);
  var baserowindex = accountrowindexes.indexOf('BaseData')+1;
  datastartindex = accountrowindexes.indexOf('Categories')+1;
  dataendindex = accountrowindexes.indexOf('Totals:TOTAL')+1;
  var startyear = parametersheet.getRange(2,2).getValue();
  var endyear = parametersheet.getRange(2,3).getValue();
  var yearindexes = actualsheet.getRange(baserowindex + ":" + baserowindex).getValues();
  var yearindexes = yearindexes[0];
  var startyearindex = yearindexes.indexOf(startyear)+1;
  var endyearindex = yearindexes.indexOf(endyear)+1;
  changesheet.clear();
  theSpreadsheet.setActiveSheet(changesheet);
  var colincrement = 4;
  var datalength = dataendindex - datastartindex + 1;
  var labeldata = actualsheet.getRange(datastartindex,1,datalength,2).getValues();
  var rowoffset = 0;
  var destcell = changesheet.getRange(2,1);
  destcell.offset(rowoffset,0).setValue('Section:Actual');
  rowoffset++;
  destcell.offset(rowoffset,0,datalength,2).setValues(labeldata);
  rowoffset+=datalength;
  destcell.offset(rowoffset,0).setValue('SectionEnd:Actual');
  rowoffset+=2;
  destcell.offset(rowoffset,0).setValue('Section:Constant');
  rowoffset++;
  destcell.offset(rowoffset,0,datalength,2).setValues(labeldata);
  rowoffset+=datalength;
  destcell.offset(rowoffset,0).setValue('SectionEnd:Constant');
  rowoffset+=2;
  destcell.offset(rowoffset,0).setValue('Section:Common');
  rowoffset++;
  destcell.offset(rowoffset,0,datalength,2).setValues(labeldata);
  rowoffset+=datalength;
  destcell.offset(rowoffset,0).setValue('SectionEnd:Common');
  rowoffset++;
  destcell.offset(rowoffset,7).setValue(''); // blank row; make sure there is room for required cols
  var changeaccountrowindexes = changesheet.getRange('A:A').getValues();
  changeaccountrowindexes = _flattenValues(changeaccountrowindexes);
  // startyear
  var coloffset = 2;
  var rowoffset = 1;
  var amountdata = actualsheet.getRange(datastartindex,startyearindex,datalength,1).getValues();
  destcell.offset(rowoffset,coloffset,datalength,1).setValues(amountdata);
  rowoffset+=datalength+3;
  destcell.offset(rowoffset,coloffset,datalength,1).setValues(amountdata);
  rowoffset+=datalength+3;
  destcell.offset(rowoffset,coloffset,datalength,1).setValues(amountdata);
  // endyear
  var coloffset = 3;
  var rowoffset = 1;
  var amountdata = actualsheet.getRange(datastartindex,endyearindex,datalength,1).getValues();
  destcell.offset(rowoffset,coloffset,datalength,1).setValues(amountdata);
  rowoffset+=datalength+3;
  destcell.offset(rowoffset,coloffset,datalength,1).setValues(amountdata);
  rowoffset+=datalength+3;
  destcell.offset(rowoffset,coloffset,datalength,1).setValues(amountdata);


  /* -------------------------------------------------------
   * create data for each dataset - actual, constant, common
   * ------------------------------------------------------- */
  // get supporting information
  var metarange = parametersheet.getRange('A:A');
  var metaindexes = metarange.getValues();  
  var metaindexes = _flattenValues(metaindexes);
  var categorytableindex = metaindexes.indexOf('CategoryTable') +1;
  var categorytableendindex = metaindexes.indexOf('CategoryTableEnd')+1;
  var classtableindex = metaindexes.indexOf('DomainTable')+1;
  var classtableendindex = metaindexes.indexOf('DomainTableEnd')+1;
  var totaltableindex = metaindexes.indexOf('TotalData')+1;
  var totaltableendindex = metaindexes.indexOf('TotalData')+1;
  // get category table
  var categoryCodes = parametersheet.getRange(
    categorytableindex, // row
    3, // col
    categorytableendindex - categorytableindex +1 // length
  ).getValues();
  categoryCodes = _flattenValues(categoryCodes);
  // get class table
  var classCodes = parametersheet.getRange(
    classtableindex, // row
    3, // col
    classtableendindex - classtableindex + 1 // length
  ).getValues();
  classCodes = _flattenValues(classCodes);
  // get total code
  var totalCodes = parametersheet.getRange(
    totaltableindex, // row
    3, // col
    totaltableendindex - totaltableindex +1 // length, not including table column headers
  ).getValues();
  totalCodes = _flattenValues(totalCodes);
  var inflationtablestart = metaindexes.indexOf('CPI adjustment year')+1+1;
  var inflationtableend = metaindexes.indexOf('CPI adjustment');
  var inflationtable = parametersheet.getRange(inflationtablestart+":"+inflationtableend).getValues();
  var inflationyearsindexes = inflationtable[0];
  var inflationdata = inflationtable[1];
  var inflationstartindex = inflationyearsindexes.indexOf(startyear);
  var inflationendindex = inflationyearsindexes.indexOf(endyear);
  // the required inflation numbers, for application to constant section
  var inflationstart = inflationdata[inflationstartindex];
  var inflationend = inflationdata[inflationendindex];
  // ========================================
  // transform numbers and calculate changes
  // ========================================
  destcell = changesheet.getRange(1,1);
  // transform Actual section
  var sectionstartindex = changeaccountrowindexes.indexOf('Section:Actual');
  var sectionoffset = sectionstartindex;
  var rowoffset = sectionoffset;
  var sectionendindex = changeaccountrowindexes.indexOf('SectionEnd:Actual');
  var sectionrowindexes = changeaccountrowindexes.slice(sectionstartindex,sectionendindex+1);
  var metasection = 'Actual';
  var callback = null;
  var parms = {};
  var metacode = 'Category';
  _calcChanges(metasection,metacode,categoryCodes,sectionrowindexes,destcell,rowoffset,callback,parms);
  metacode = 'Domain';
  _calcChanges(metasection,metacode,classCodes,sectionrowindexes,destcell,rowoffset,callback,parms);
  metacode = 'Domain Totals';
  _calcChanges(metasection,metacode,totalCodes,sectionrowindexes,destcell,rowoffset,callback,parms);
  // transform Constant section
  sectionstartindex = changeaccountrowindexes.indexOf('Section:Constant');
  sectionoffset = sectionstartindex;
  rowoffset = sectionoffset;
  sectionendindex = changeaccountrowindexes.indexOf('SectionEnd:Constant');
  sectionrowindexes = changeaccountrowindexes.slice(sectionstartindex,sectionendindex+1);
  metasection = 'Constant';
  callback = _applyInflation;
  parms = {};
  parms.inflationstart = inflationstart;
  parms.inflationend = inflationend;
  metacode = 'Category';
  _calcChanges(metasection,metacode,categoryCodes,sectionrowindexes,destcell,rowoffset,callback,parms);
  metacode = 'Domain';
  _calcChanges(metasection,metacode,classCodes,sectionrowindexes,destcell,rowoffset,callback,parms);
  metacode = 'Domain Totals';
  _calcChanges(metasection,metacode,totalCodes,sectionrowindexes,destcell,rowoffset,callback,parms);
  // transform Common section
  sectionstartindex = changeaccountrowindexes.indexOf('Section:Common');
  sectionoffset = sectionstartindex;
  rowoffset = sectionoffset;
  sectionendindex = changeaccountrowindexes.indexOf('SectionEnd:Common');
  sectionrowindexes = changeaccountrowindexes.slice(sectionstartindex,sectionendindex+1);
  metasection = 'Common';
  callback = _convertToCommon;
  parms = {};
  metacode = 'Category';
  _calcChanges(metasection,metacode,categoryCodes,sectionrowindexes,destcell,rowoffset,callback,parms);
  metacode = 'Domain';
  _calcChanges(metasection,metacode,classCodes,sectionrowindexes,destcell,rowoffset,callback,parms);
  metacode = 'Domain Totals';
  _calcChanges(metasection,metacode,totalCodes,sectionrowindexes,destcell,rowoffset,callback,parms);
}
var _applyInflation = function(row,parms) {
  row[2] *= parms.inflationstart;
  row[3] *= parms.inflationend;
  return row;
}


var _convertToCommon = function(workdata) {
  var totalrow = workdata[workdata.length -1];
  var starttotal = totalrow[2];
  var endtotal = totalrow[3]; 
  for (var i = 1; i < workdata.length; i++) {
    var row = workdata[i];
    row[2] = (row[2]/starttotal) * 100;
    row[3] = (row[3]/endtotal) * 100;
    workdata[i] = row;
  }
  return workdata;
}
function _calcChanges(metasection,metacode,codes,sectionrowindexes,destcell,rowoffset,callback,parms) {
  for each (var code in codes) {
    var startoffset = sectionrowindexes.indexOf(metacode+':'+code);
    var endoffset = sectionrowindexes.indexOf('Totals:' + code);
    var workdata = destcell.offset(rowoffset+startoffset,0,endoffset-startoffset+1,8).getValues();
    if (callback && (metasection == 'Common')) {
      callback(workdata);
    }
    var titlerow = workdata[0];
    titlerow[4] = 'change ('+metasection+')';
    titlerow[5] = '% of start';
    titlerow[6] = '% of end'
    titlerow[7] = '% of total change';
    workdata[0] = titlerow;
    var totalrow = workdata[workdata.length - 1];
    var newarray = [];
    for each (var item in totalrow) {
      newarray.push(item);
    }
    totalrow = newarray;
    if (callback && (metasection == 'Constant')) {
      totalrow = callback(totalrow,parms);
    }
    totalrow[4] = totalrow[3] - totalrow[2];
    for (var i = 1; i < workdata.length; i++) {
      var row = workdata[i];
      if (callback && (metasection == 'Constant')) {
        row = callback(row,parms);
      }
      var is_startval = Boolean(row[2]);
      // calc change
      if (is_startval) {
        row[4] = row[3] - row[2];
      } else {
        row[4] = row[3];
      }
      // calc % of start
      if (is_startval) {
        row[5] = parseFloat(((row[4]/row[2]) * 100).toFixed(1));
      } else {
        row[5] = 'n/a';
      }
      // calc % of end
      row[6] = parseFloat(((row[4]/row[3]) * 100).toFixed(1));
      // calc % of total
      if (totalrow[4] == 0) {
        row[7] = 'n/a';
      } else {
        row[7] = parseFloat(((row[4]/totalrow[4]) * 100).toFixed(1));
      }
      workdata[i] = row;
    }
    destcell.offset(rowoffset+startoffset,0,endoffset-startoffset+1,8).setValues(workdata);
  }
}


function _callActualTrendCharts(theSpreadsheet) {
  _setTrendCharts(theSpreadsheet, 'Actual Trend Charts', 'Actual Trend Charts Transposed',null);
}


function _callConstantTrendCharts(theSpreadsheet) {
  _setTrendCharts(theSpreadsheet, 'Constant Trend Charts', 'Constant Trend Charts Transposed',_transformtoconstant);
}


function _callCommonTrendCharts(theSpreadsheet) {
  _setTrendCharts(theSpreadsheet, 'Common Trend Charts', 'Common Trend Charts Transposed',_transformtocommon);
}


var _transformtoconstant = function(writedata,parms) {
//  Logger.log(writedata);
  var startyear = parms.startyear;
  var endyear = parms.endyear;
  var inflationdata = parms.inflationdata;
  for (var i = 1; i < writedata.length; i++) {
    var row = writedata[i];
    for (var j = 2; j < row.length; j++) {
      row[j] = row[j] * inflationdata[j-2];
    }
    writedata[i] = row;
  }
  return writedata;
}


var _transformtocommon = function(writedata,parms) {
  var totalsdata = writedata[writedata.length - 1];
  Logger.log(totalsdata);
  for (var i = 1; i < (writedata.length - 1); i++) {
    var row = writedata[i];
    for (var j = 2; j < row.length; j++) {
      row[j] = (row[j]/totalsdata[j]) * 100;
    }
    writedata[i] = row;
  }
  for (var j = 2; j < totalsdata.length; j++) {
      totalsdata[j] = 100;
  }
  writedata[writedata.length - 1] = totalsdata;
  Logger.log(writedata);
  return writedata;
}


function _setTrendCharts(theSpreadsheet,chartsheetname,chartscratchsheetname,transformcallback) {
  var theSpreadsheet = theSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  var actualsheet, chartsheet, chartscratchsheet, parametersheet;
  actualsheet = theSpreadsheet.getSheetByName('Consolidated Actual');
  chartsheet = theSpreadsheet.getSheetByName(chartsheetname);
  chartscratchsheet = theSpreadsheet.getSheetByName(chartscratchsheetname);
  parametersheet = theSpreadsheet.getSheetByName('Parameters');
  var accountrowindexes, refyear, titleindexes;
  var datastartindex, dataendindex, refyearcolindex;
  accountrowindexes = actualsheet.getRange('A:A').getValues();
  accountrowindexes = _flattenValues(accountrowindexes);
  var baserowindex = accountrowindexes.indexOf('BaseData')+1;
  datastartindex = accountrowindexes.indexOf('Categories')+1;
  dataendindex = accountrowindexes.indexOf('Totals:TOTAL')+1;
  var refyear = parametersheet.getRange(2,3).getValue();
  var yearindexes = actualsheet.getRange(baserowindex + ":" + baserowindex).getValues();
  var yearindexes = yearindexes[0];
  var yearindex = yearindexes.indexOf(refyear)+1;
  chartsheet.clear();
  var charts = chartsheet.getCharts();
  for each (var chart in charts) {
    chartsheet.removeChart(chart);
  }
  theSpreadsheet.setActiveSheet(chartsheet);
  var trenddata = actualsheet.getRange(datastartindex,1,dataendindex - datastartindex + 1,yearindex).getValues();
  // **TODO collection of code tables should be done by utility function, then back-fitted throughout.
  // assumption: parameter table order is identical to actualsheet order
  // get category table
  var metarange = parametersheet.getRange('A:A');
  var metaindexes = metarange.getValues();  
  var metaindexes = _flattenValues(metaindexes);
  var categorytableindex = metaindexes.indexOf('CategoryTable') +1;
  var categorytableendindex = metaindexes.indexOf('CategoryTableEnd')+1;
  var classtableindex = metaindexes.indexOf('DomainTable')+1;
  var classtableendindex = metaindexes.indexOf('DomainTableEnd')+1;
  var totaltableindex = metaindexes.indexOf('TotalData')+1;
  var totaltableendindex = metaindexes.indexOf('TotalData')+1;
  var categoryCodes = parametersheet.getRange(
    categorytableindex, // row
    3, // col
    categorytableendindex - categorytableindex +1 // length
  ).getValues();
  categoryCodes = _flattenValues(categoryCodes);
  // get class table
  var classCodes = parametersheet.getRange(
    classtableindex, // row
    3, // col
    classtableendindex - classtableindex + 1 // length
  ).getValues();
  classCodes = _flattenValues(classCodes);
  // get total code
  var totalCodes = parametersheet.getRange(
    totaltableindex, // row
    3, // col
    totaltableendindex - totaltableindex +1 // length, not including table column headers
  ).getValues();
  totalCodes = _flattenValues(totalCodes);
  // insert space for charts
  var rowoffset = 0;
  var rowscountforchart = 18;
  var slicestart = 0;
  var sliceend, slicemarker,writedata;
  var destcell, rowlength;
  accountrowindexes.splice(0,datastartindex - 1);
  // write header
  // write categories header
  slicemarker = accountrowindexes.indexOf('Categories') + 1;
  sliceend = slicemarker;
  writedata = trenddata.slice(slicestart,sliceend);
  destcell = chartsheet.getRange(2,1);
  rowlength = sliceend - slicestart;
  destcell.offset(rowoffset,0,rowlength,yearindex).setValues(writedata);
  rowoffset += rowlength;
  // write categories
  slicestart = sliceend + 1;
  var transformparm = {};
  if (chartsheetname == 'Constant Trend Charts') {
    transformparm.startyear = parametersheet.getRange(2,2).getValue();
    transformparm.endyear = refyear;
    var inflationtablestart = metaindexes.indexOf('CPI adjustment year')+1+1;
    var inflationtableend = metaindexes.indexOf('CPI adjustment');
    var inflationtable = parametersheet.getRange(inflationtablestart+":"+inflationtableend).getValues();
    var inflationyearsindexes = inflationtable[0];
    var inflationdata = inflationtable[1];
    var inflationstartindex = inflationyearsindexes.indexOf(transformparm.startyear);
    var inflationendindex = inflationyearsindexes.indexOf(transformparm.endyear);
    transformparm.inflationdata = inflationdata.slice(inflationstartindex,inflationendindex+1);
  }
  for each (var category in categoryCodes) {
    destcell.offset(rowoffset,0,1,1).setValue(''); // blank row
    rowoffset++;
    destcell.offset(rowoffset,0,1,1).setValue('Category Chart:' + category);
    rowoffset++;
    destcell.offset(rowoffset,0,rowscountforchart,1).setValue(''); // space for trend chart
    rowoffset+=rowscountforchart;
    slicestart = accountrowindexes.indexOf('Category:'+category);
    sliceend = accountrowindexes.indexOf('Totals:'+category) + 1;
    writedata = trenddata.slice(slicestart,sliceend);
    if (transformcallback) {
      writedata = transformcallback(writedata,transformparm);
    }
    rowlength = sliceend - slicestart;
    destcell.offset(rowoffset,0,rowlength,yearindex).setValues(writedata);
    rowoffset += rowlength;   
  }
  // write classes header
  rowoffset++;
  slicemarker = accountrowindexes.indexOf('Domains') + 1;
  slicestart = slicemarker - 1;
  sliceend = slicemarker;
  writedata = trenddata.slice(slicestart,sliceend);
  rowlength = sliceend - slicestart;
  destcell.offset(rowoffset,0,rowlength,yearindex).setValues(writedata);
  rowoffset += rowlength;
  // write classes
  slicestart = sliceend + 1;
  for each (var class in classCodes) {
    destcell.offset(rowoffset,0,1,1).setValue(''); // blank row
    rowoffset++;
    destcell.offset(rowoffset,0,1,1).setValue('Domain Chart:' + class);
    rowoffset++;
    destcell.offset(rowoffset,0,rowscountforchart,1).setValue(''); // space for trend chart
    rowoffset+=rowscountforchart;
    slicestart = accountrowindexes.indexOf('Domain:'+class);
    sliceend = accountrowindexes.indexOf('Totals:'+class) + 1;
    writedata = trenddata.slice(slicestart,sliceend);
    if (transformcallback) {
      writedata = transformcallback(writedata,transformparm);
    }
    rowlength = sliceend - slicestart;
    destcell.offset(rowoffset,0,rowlength,yearindex).setValues(writedata);
    rowoffset += rowlength;    
  }
  // write totals header
  rowoffset++;
  slicemarker = accountrowindexes.indexOf('Domain Totals') + 1;
  slicestart = slicemarker - 1;
  sliceend = slicemarker;
  writedata = trenddata.slice(slicestart,sliceend);
  rowlength = sliceend - slicestart;
  destcell.offset(rowoffset,0,rowlength,yearindex).setValues(writedata);
  rowoffset += rowlength;
  // write totals
  slicestart = sliceend + 1;
  for each (var total in totalCodes) {
    destcell.offset(rowoffset,0,1,1).setValue(''); // blank row
    rowoffset++;
    destcell.offset(rowoffset,0,1,1).setValue('Domain Totals Chart:' + total);
    rowoffset++;
    destcell.offset(rowoffset,0,rowscountforchart,1).setValue(''); // space for trend chart
    rowoffset+=rowscountforchart;
    slicestart = accountrowindexes.indexOf('Domain Totals:'+total);
    sliceend = accountrowindexes.indexOf('Totals:'+total) + 1;
    writedata = trenddata.slice(slicestart,sliceend);
    if (transformcallback) {
      writedata = transformcallback(writedata,transformparm);
    }
    rowlength = sliceend - slicestart;
    destcell.offset(rowoffset,0,rowlength,yearindex).setValues(writedata);
    rowoffset += rowlength;    
  }
  // write footer to sheet
  rowoffset+=2;
  chartsheet.getRange(rowoffset,1).setValue(''); // empty row after data for chart
  var trendaccountrowindexes = chartsheet.getRange('A:A').getValues();
  trendaccountrowindexes = _flattenValues(trendaccountrowindexes);
  // write trend charts
  // use deleteRows and getMaxRows() to reset scratch sheet, for transposed tables
  if (chartscratchsheet.getMaxRows() > 1) {
    chartscratchsheet.deleteRows(2,chartscratchsheet.getMaxRows()-1);
  }
  var theChartHeight = 360;
  var theChartWidth = 900;
  var theDataWidth = yearindex - 1;
  for each (var category in categoryCodes) {
    _createTrendChart('Category',category,trendaccountrowindexes,chartsheet,chartscratchsheet,theChartHeight, theChartWidth, theDataWidth);
  }  
  for each (var class in classCodes) {
    _createTrendChart('Domain',class,trendaccountrowindexes,chartsheet,chartscratchsheet,theChartHeight, theChartWidth, theDataWidth);
  }  
  for each (var total in totalCodes) {
    _createTrendChart('Domain Totals',total,trendaccountrowindexes,chartsheet,chartscratchsheet,theChartHeight, theChartWidth, theDataWidth);
  }  
}


function _createTrendChart(metacode,code,trendaccountrowindexes,sheet,chartscratchsheet,theChartHeight,theChartWidth,theDataWidth) {
  var startrowindex = trendaccountrowindexes.indexOf(metacode+':'+code)+1;
  var endrowindex = trendaccountrowindexes.indexOf('Totals:'+code)+1;
  var thechartpos = trendaccountrowindexes.indexOf(metacode+' Chart:'+code)+1;
//  var amount = sheet.getRange(endrowindex,3).getValue();
//  var year = sheet.getRange(startrowindex,3).getValue();
  var nameindex = trendaccountrowindexes.indexOf(metacode + ':' + code) + 1;
  var sheetname = sheet.getName();
  var codename = sheet.getRange(nameindex,2).getValue();
  var title = 'Budget trend chart - ' + metacode + ': ' + codename;
  if (sheetname == 'Common Trend Charts') {
    title += " (% of tota), by year";
  } else if (sheetname == 'Actual Trend Charts') {
    title += " (000's), nominal dollars by year";
  } else if (sheetname == 'Constant Trend Charts') {
    title += " (000's), inflation adjusted dollars by year";
  }
var theHeight = endrowindex - startrowindex;
  var transposed = _transpose(sheet.getRange(startrowindex,2,theHeight,theDataWidth).getValues());
  for (var i = 0; i < transposed.length; i++) {
    var r = transposed[i];
    for (var j = 0; j < r.length; j++) {
      if (!r[j]) r[j] = 0;
    }
    transposed[i] = r;
  }
// write the transposed data to a hidden scratch sheet
  var scratchrowindex = chartscratchsheet.getMaxRows();
  scratchrowindex++;
  chartscratchsheet.getRange(scratchrowindex,1).setValue('Transposed:'+ code);
  scratchrowindex++;
  var sourceheight = transposed.length;
  var sourcewidth = transposed[0].length;
  chartscratchsheet.getRange(scratchrowindex,1,sourceheight,sourcewidth).setValues(transposed);
  var sourcestartindex = scratchrowindex;
  scratchrowindex+=sourceheight;
  chartscratchsheet.getRange(scratchrowindex,1).setValue('');  
  var chart = sheet.newChart()
     .setChartType(Charts.ChartType.AREA)
     .addRange(chartscratchsheet.getRange(sourcestartindex,1,transposed.length,transposed[0].length))
//     .setDataTable(data) -- does not work with embedded charts
     .setPosition(thechartpos+1, 1, 0, 0)
     .setOption('title',title)
     .setOption('height',theChartHeight)
     .setOption('width',theChartWidth)
     .asAreaChart()
     .setStacked()
//     .reverseCategories()
     .setOption('useFirstColumnAsDomain',true)
     .setOption('legend.textStyle',{fontSize:10})
     .build();
  sheet.insertChart(chart);
}


function _setActualRefCharts(theSpreadsheet) {
  var theSpreadsheet = theSpreadsheet || SpreadsheetApp.getActiveSpreadsheet();
  /* ------------------------------------------
   * copy data to 'Actual Reference Charts' sheet
   * ------------------------------------------ */
  var actualsheet, refchartsheet, parametersheet;
  actualsheet = theSpreadsheet.getSheetByName('Consolidated Actual');
  refchartsheet = theSpreadsheet.getSheetByName('Actual Reference Charts');
  parametersheet = theSpreadsheet.getSheetByName('Parameters');
  var accountrowindexes, refyear, titleindexes;
  var datastartindex, dataendindex, refyearcolindex;
  accountrowindexes = actualsheet.getRange('A:A').getValues();
  accountrowindexes = _flattenValues(accountrowindexes);
  var baserowindex = accountrowindexes.indexOf('BaseData')+1;
  datastartindex = accountrowindexes.indexOf('Categories')+1;
  dataendindex = accountrowindexes.indexOf('Totals:TOTAL')+1;
  var refyear = parametersheet.getRange(2,3).getValue();
  var yearindexes = actualsheet.getRange(baserowindex + ":" + baserowindex).getValues();
  var yearindexes = yearindexes[0];
  var yearindex = yearindexes.indexOf(refyear)+1;
  refchartsheet.clear();
  var charts = refchartsheet.getCharts();
  for each (var chart in charts) {
    refchartsheet.removeChart(chart);
  }
  theSpreadsheet.setActiveSheet(refchartsheet);
  var labeldata = actualsheet.getRange(datastartindex,1,dataendindex - datastartindex + 1,2).getValues();
  var destcell = refchartsheet.getRange(2,1);
  destcell.offset(0,0,dataendindex - datastartindex + 1,2).setValues(labeldata);
  var amountdata = actualsheet.getRange(datastartindex,yearindex,dataendindex - datastartindex + 1,1).getValues();
  var destcell = refchartsheet.getRange(2,3);
  destcell.offset(0,0,dataendindex - datastartindex + 1,1).setValues(amountdata);
  refchartsheet.getRange(dataendindex - datastartindex + 3,1).setValue(''); // empty row after data for chart
  var refaccountrowindexes = refchartsheet.getRange('A:A').getValues();
  refaccountrowindexes = _flattenValues(refaccountrowindexes);
  /* ------------------------------------------
   * create chart for each data table
   * ------------------------------------------ */
  var metarange = parametersheet.getRange('A:A');
  var metaindexes = metarange.getValues();  
  var metaindexes = _flattenValues(metaindexes);
  var categorytableindex = metaindexes.indexOf('CategoryTable') +1;
  var categorytableendindex = metaindexes.indexOf('CategoryTableEnd')+1;
  var classtableindex = metaindexes.indexOf('DomainTable')+1;
  var classtableendindex = metaindexes.indexOf('DomainTableEnd')+1;
  var totaltableindex = metaindexes.indexOf('TotalData')+1;
  var totaltableendindex = metaindexes.indexOf('TotalData')+1;
  // get category table
  var categoryCodes = parametersheet.getRange(
    categorytableindex, // row
    3, // col
    categorytableendindex - categorytableindex +1 // length
  ).getValues();
  categoryCodes = _flattenValues(categoryCodes);
  // create category base charts
  for each (var category in categoryCodes) {
    _createCategoryRefChart(category,refaccountrowindexes,refchartsheet);
  }
  // get class table
  var classCodes = parametersheet.getRange(
    classtableindex, // row
    3, // col
    classtableendindex - classtableindex + 1 // length
  ).getValues();
  classCodes = _flattenValues(classCodes);
  // create class base charts
  for each (var class in classCodes) {
    _createClassRefChart(class,refaccountrowindexes,refchartsheet);
  }
  // get total code
  var totalCodes = parametersheet.getRange(
    totaltableindex, // row
    3, // col
    totaltableendindex - totaltableindex +1 // length, not including table column headers
  ).getValues();
  totalCodes = _flattenValues(totalCodes);
  // create total base chart
  for each (var total in totalCodes) {
    _createTotalRefChart(total,refaccountrowindexes,refchartsheet);
  }
  /* ------------------------------------------
   * create descriptive titles for domain groupings
   * ------------------------------------------ */
  // ** TODO ** the following math needs to be generalized - it currently makes assumptions about orders of magnitude
  var totalsstartindex = refaccountrowindexes.indexOf('Domain Totals')+1;
  var totalsendindex = refaccountrowindexes.indexOf('Totals:TOTAL')+1;
  var budgettotal = refchartsheet.getRange(totalsendindex,3).getValue();
  var budgettotalstring = budgettotal/100000;
  budgettotalstring = Math.round(budgettotalstring);
  budgettotalstring = budgettotalstring/10;
  budgettotalstring = '$' + String(budgettotalstring).substr(0,4) + 'B';
  for each (class in classCodes) {
    var classindex = refaccountrowindexes.indexOf(class)+1;
    var classtotal = refchartsheet.getRange(classindex,3).getValue();
    var classpercent = (classtotal/budgettotal) * 1000;
    classpercent = Math.round(classpercent);
    classpercent = classpercent/10;
    classpercent = String(classpercent).substr(0,4) + '%';
    classtotal = classtotal/100000;
    classtotal = Math.round(classtotal);
    classtotal = classtotal/10;
    classtotal = '$' + String(classtotal).substr(0,3) + 'B';
    var classmarkerindex = refaccountrowindexes.indexOf('Marker:'+class)+1;
    var classtitle = refchartsheet.getRange(classmarkerindex,2).getValue();
    classtitle += ': The following group of categories represents ' + 
      classtotal + ' or ' + classpercent + ' of the overall ' + 
      budgettotalstring + ' budget for ' + refyear;
    refchartsheet.getRange(classmarkerindex,2).setValue(classtitle);
  }
}


function _createCategoryRefChart(category,refaccountrowindexes,sheet) {
  var startrowindex = refaccountrowindexes.indexOf('Category:'+category)+1;
  var endrowindex = refaccountrowindexes.indexOf('Totals:'+category)+1;
  var amount = sheet.getRange(endrowindex,3).getValue();
  var year = sheet.getRange(startrowindex,3).getValue();
  var title = year + " " + sheet.getRange(startrowindex,2).getValue() + ": $" + _CurrencyFormat(amount) + " (000's)";
  var theHeight = endrowindex - startrowindex - 1;
  var theChartHeight = (theHeight * 20) + (3 * 20);
  var chart = sheet.newChart()
     .setChartType(Charts.ChartType.PIE)
     .addRange(sheet.getRange(startrowindex + 1,2,theHeight,2))
     .setPosition(startrowindex, 4, 0, 0)
     .setOption('title',title)
     .setOption('height',theChartHeight)
     .setOption('legend',{position: 'labeled'})
     .build();
  sheet.insertChart(chart);
}


function _createClassRefChart(class,refaccountrowindexes,sheet) {
  var startrowindex = refaccountrowindexes.indexOf('Domain:'+class)+1;
  var endrowindex = refaccountrowindexes.indexOf('Totals:'+class)+1;
  var amount = sheet.getRange(endrowindex,3).getValue();
  var year = sheet.getRange(startrowindex,3).getValue();
  var title = year + " " + sheet.getRange(startrowindex,2).getValue() + ": $" + _CurrencyFormat(amount) + " (000's)";
  var theHeight = endrowindex - startrowindex - 1;
  var theChartHeight = (theHeight * 20) + (3 * 20);
  var chart = sheet.newChart()
     .setChartType(Charts.ChartType.PIE)
     .addRange(sheet.getRange(startrowindex + 1,2,theHeight,2))
     .setPosition(startrowindex, 4, 0, 0)
     .setOption('title',title)
     .setOption('height',theChartHeight)
     .setOption('legend',{position: 'labeled'})
     .build();
  sheet.insertChart(chart);
}


function _createTotalRefChart(total,refaccountrowindexes,sheet) {
  var startrowindex = refaccountrowindexes.indexOf('Domain Totals:'+total)+1;
  var endrowindex = refaccountrowindexes.indexOf('Totals:'+total)+1;
  var amount = sheet.getRange(endrowindex,3).getValue();
  var year = sheet.getRange(startrowindex,3).getValue();
  var title = year + " " + sheet.getRange(startrowindex,2).getValue() + ": $" + _CurrencyFormat(amount) + " (000's)";
  var theHeight = endrowindex - startrowindex - 1;
  var theChartHeight = (theHeight * 20) + (3 * 20);
  var chart = sheet.newChart()
     .setChartType(Charts.ChartType.PIE)
     .addRange(sheet.getRange(startrowindex + 1,2,theHeight,2))
     .setPosition(startrowindex, 4, 0, 0)
     .setOption('title',title)
     .setOption('height',theChartHeight)
     .setOption('legend',{position: 'labeled'})
     .build();
  sheet.insertChart(chart);
}


function _setBaseAggregates(theSpreadsheetparm) {
  var theSpreadsheet = theSpreadsheetparm || SpreadsheetApp.getActiveSpreadsheet();
  var actualsheet, parametersheet;
  var metarange, metaindexes,yearsindex;
  var accounttableindex,accounttableendindex;
  var startYear, endYear;
  var indexTableHeadings, referenceYearColumnIndex, accountCodes, rollupColumnIndex, rollupCodes, tablelength;
  var accountrowanchorcell;
  actualsheet = theSpreadsheet.getSheetByName('Consolidated Actual');
  theSpreadsheet.setActiveSheet(actualsheet);
  // collect various indexes from parameter sheet
  parametersheet = theSpreadsheet.getSheetByName('Parameters');
  var categoryTableDictionary = _getClassificationDictionary(theSpreadsheet,'CategoryTable','CategoryTableEnd');
  var classTableDictionary = _getClassificationDictionary(theSpreadsheet,'DomainTable','DomainTableEnd');
  var totalTableDictionary = _getClassificationDictionary(theSpreadsheet,'TotalData','TotalData');
  metarange = parametersheet.getRange('A:A');
  metaindexes = metarange.getValues();  
  metaindexes = _flattenValues(metaindexes);
  yearsindex = metaindexes.indexOf('Range')+1;
  accounttableindex = metaindexes.indexOf('IndexTable')+1;
  accounttableendindex = metaindexes.indexOf('IndexTableEnd')+1;
  var categorytableindex = metaindexes.indexOf('CategoryTable') +1;
  var categorytableendindex = metaindexes.indexOf('CategoryTableEnd')+1;
  var classtableindex = metaindexes.indexOf('DomainTable')+1;
  var classtableendindex = metaindexes.indexOf('DomainTableEnd')+1;
  var totaltableindex = metaindexes.indexOf('TotalData')+1;
  var totaltableendindex = metaindexes.indexOf('TotalData')+1;
  startYear = Math.round(parametersheet.getRange(yearsindex,2).getValue());
  endYear = Math.round(parametersheet.getRange(yearsindex,3).getValue());
  // collect table headings and reference year index from parameter sheet
  indexTableHeadings = parametersheet.getRange(accounttableindex + ":" + accounttableindex).getValues();
  indexTableHeadings = indexTableHeadings[0]; // get row values;
  var tablewidth = endYear - startYear + 1 + 2; // +1 for offset; +2 for code/title
  // get base account codes and their rollup codes
  referenceYearColumnIndex = indexTableHeadings.indexOf(endYear)+1;
  accountCodes = parametersheet.getRange(
    accounttableindex + 1, // row
    referenceYearColumnIndex, // col
    accounttableendindex - accounttableindex // length, not including table column headers
  ).getValues();
  accountCodes = _flattenValues(accountCodes);
  accountCodes = _compactArray(accountCodes);
  tablelength = accountCodes.length;
  rollupColumnIndex = indexTableHeadings.indexOf('RollUpto')+1;
  rollupCodes = parametersheet.getRange(
    accounttableindex + 1, // row
    rollupColumnIndex, // col
    accounttableendindex - accounttableindex // length, not including table column headers
  ).getValues();
  rollupCodes = _flattenValues(rollupCodes);
  rollupCodes = _compactArray(rollupCodes);
  // get category account codes and their rollup codes
  var categoryCodes = parametersheet.getRange(
    categorytableindex, // row
    3, // col
    categorytableendindex - categorytableindex +1 // length
  ).getValues();
  categoryCodes = _flattenValues(categoryCodes);
  var categorytablelength = categoryCodes.length;
  var categoryRollupCodes = parametersheet.getRange(
    categorytableindex, // row
    2, // col
    categorytableendindex - categorytableindex +1 // length
  ).getValues();
  categoryRollupCodes = _flattenValues(categoryRollupCodes);
  // get class account codes and their rollup codes
  var classCodes = parametersheet.getRange(
    classtableindex, // row
    3, // col
    classtableendindex - classtableindex + 1 // length
  ).getValues();
  classCodes = _flattenValues(classCodes);
  var classtablelength = classCodes.length;
  var classRollupCodes = parametersheet.getRange(
    classtableindex, // row
    2, // col
    classtableendindex - classtableindex +1 // length
  ).getValues();
  classRollupCodes = _flattenValues(classRollupCodes);
  // get total account codes
  var totalCodes = parametersheet.getRange(
    totaltableindex, // row
    3, // col
    totaltableendindex - totaltableindex +1 // length, not including table column headers
  ).getValues();
  totalCodes = _flattenValues(totalCodes);
  // get base account row data
  var actualtitlerowoffset = 1;
  var basetablerowindex = actualtitlerowoffset+2;
  accountrowanchorcell = actualsheet.getRange(basetablerowindex,1);//offset + 1 for index, + 1 for 1st past title
  // collect base account values in category rollup arrays
  var rollupcollection = {};
  for (var offset = 0; offset < tablelength; offset++) {
    var rollupcode = rollupCodes[offset];
    if (rollupcollection[rollupcode] == undefined) {
      rollupcollection[rollupcode] = [];
    }
    var rowvalues = accountrowanchorcell.offset(offset,0,1,tablewidth).getValues();
    rollupcollection[rollupcode].push(rowvalues);
  }
  // ==============================[ CATEGORIES ]================================
  // write values to rollup groups
  var markercode = '';
  var rolluparearowindex = basetablerowindex+tablelength;
  var rollupgroupanchorcell = actualsheet.getRange(rolluparearowindex,1);
  var offset = 0;
  var refcell = actualsheet.getRange(rolluparearowindex + offset+1,1);
  refcell.setValue('Categories');
  offset++;
  var categorysummaries = {};
  for each (var categorycode in categoryCodes) {
    offset+=2;
    var tablepos = categoryCodes.indexOf(categorycode);
    var rollupcode = categoryRollupCodes[tablepos];
    if (rollupcode != markercode) {
      markercode = rollupcode;
      refcell = actualsheet.getRange(rolluparearowindex + offset,1);
      refcell.setValue('Marker:'+markercode);
      refcell.offset(0,1).setValue(classTableDictionary[markercode]); // class title
      offset+=2;
    }
    refcell = actualsheet.getRange(rolluparearowindex + offset,1);
    refcell.setValue('Category:'+categorycode);
    refcell.offset(0,1).setValue(categoryTableDictionary[categorycode]); // category title
    // add year titles
    var coloffset = 2;
    for (var year = startYear; year <= endYear; year++) {
      refcell.offset(0,coloffset).setValue(year);
      coloffset++;
    }
    var groupoffset = 1;
    var categorydata = rollupcollection[categorycode];
    var categorytotals = [];
    for each (var row in categorydata) {
      var rowdata = row[0];
      refcell.offset(groupoffset,0,1,tablewidth).setValues(row);
      for (var dataoffset = 2;dataoffset < rowdata.length; dataoffset++) {
        if (groupoffset == 1) {
          categorytotals[dataoffset] = Number(rowdata[dataoffset]);
        } else {
          categorytotals[dataoffset] += Number(rowdata[dataoffset]);
        }
      }
      groupoffset++;
      offset++;
    }
    for (var dataoffset = 2;dataoffset < categorytotals.length; dataoffset++) {
      categorytotals[dataoffset] = parseFloat(categorytotals[dataoffset].toFixed(1));
    }
    // set code and description for subtotals
    categorytotals[0] = 'Totals:'+categorycode;
    categorytotals[1] = 'Totals';
    refcell.offset(groupoffset,0,1,tablewidth).setValues([categorytotals]);
    offset++;
    // reset code and description for subsequent work
    categorytotals[0] = categorycode;
    categorytotals[1] = categoryTableDictionary[categorycode];
    categorysummaries[categorycode] = categorytotals;
  }
  // aggregate groups
  // ==============================[ CLASSES ]================================
  // collect category account values into class rollup arrays
  var rollupcollection = {};
  // use categorysummaries
  var classsummaries = {};
  // initialize rollupcollection for correct order
  for each (var class in classCodes) {
    rollupcollection[class] = [];
  }
  categorytablelength = categorysummaries.length;
  for (var category in categorysummaries) {
    rowvalues = categorysummaries[category];
    var tablepos = categoryCodes.indexOf(category);
    var rollupcode = categoryRollupCodes[tablepos];
    rollupcollection[rollupcode].push(rowvalues);
  }
  // write values to rollup groups
  rolluparearowindex = rolluparearowindex + offset+1;
  refcell = actualsheet.getRange(rolluparearowindex,1);
  refcell.offset(1,0).setValue('Domains');
  var offset = 1;
  var classssummaries = {};
  for each (var classcode in classCodes) {
    offset+=2;
    refcell = actualsheet.getRange(rolluparearowindex + offset,1);
    refcell.setValue('Domain:'+classcode);
    refcell.offset(0,1).setValue(classTableDictionary[classcode]); // class title
    // add year titles
    var coloffset = 2;
    for (var year = startYear; year <= endYear; year++) {
      refcell.offset(0,coloffset).setValue(year);
      coloffset++;
    }
    var groupoffset = 1;
    var classdata = rollupcollection[classcode];
    var classtotals = [];
    for each (var row in classdata) {
      refcell.offset(groupoffset,0,1,tablewidth).setValues([row]);
      for (var dataoffset = 2;dataoffset < rowdata.length; dataoffset++) {
        if (groupoffset == 1) {
          classtotals[dataoffset] = Number(row[dataoffset]);
        } else {
          classtotals[dataoffset] += Number(row[dataoffset]);
        }
      }
      groupoffset++;
      offset++;
    }
    for (var dataoffset = 2;dataoffset < classtotals.length; dataoffset++) {
      classtotals[dataoffset] = parseFloat(classtotals[dataoffset].toFixed(1));
    }
    // set code and description for subtotals
    classtotals[0] = 'Totals:'+classcode;
    classtotals[1] = 'Totals';
    refcell.offset(groupoffset,0,1,tablewidth).setValues([classtotals]);
    offset++;
    // reset code and description for subsequent work
    classtotals[0] = classcode;
    classtotals[1] = classTableDictionary[classcode];
    classsummaries[classcode] = classtotals;
  }
  // ==============================[ TOTALS ]================================
  // collect category account values into class rollup arrays
  var rollupcollection = {};
  // use categorysummaries
  var totalsummaries = {};
  // initialize rollupcollection for correct order
  for each (var total in totalCodes) {
    rollupcollection[total] = [];
  }
  classtablelength = classsummaries.length;
  for (var class in classsummaries) {
    rowvalues = classsummaries[class];
    var tablepos = classCodes.indexOf(class);
    var rollupcode = 'TOTAL';
    rollupcollection[rollupcode].push(rowvalues);
  }
  // write values to rollup groups
  rolluparearowindex = rolluparearowindex + offset+1;
  refcell = actualsheet.getRange(rolluparearowindex,1);
  refcell.offset(1,0).setValue('Domain Totals');
  var offset = 1;
  for (var totalcode in rollupcollection) {
    offset+=2;
    refcell = actualsheet.getRange(rolluparearowindex + offset,1);
    refcell.setValue('Domain Totals:'+totalcode);
    refcell.offset(0,1).setValue(totalTableDictionary[totalcode]); // total title
    // add year titles
    var coloffset = 2;
    for (var year = startYear; year <= endYear; year++) {
      refcell.offset(0,coloffset).setValue(year);
      coloffset++;
    }
    var groupoffset = 1;
    var totaldata = rollupcollection[totalcode];
    var totaltotals = [];
    for each (var row in totaldata) {
      refcell.offset(groupoffset,0,1,tablewidth).setValues([row]);
      for (var dataoffset = 2;dataoffset < rowdata.length; dataoffset++) {
        if (groupoffset == 1) {
          totaltotals[dataoffset] = Number(row[dataoffset]);
        } else {
          totaltotals[dataoffset] += Number(row[dataoffset]);
        }
      }
      groupoffset++;
      offset++;
    }
    for (var dataoffset = 2;dataoffset < totaltotals.length; dataoffset++) {
      totaltotals[dataoffset] = parseFloat(totaltotals[dataoffset].toFixed(1));
    }
    // set code and description for subtotals
    totaltotals[0] = 'Totals:'+totalcode;
    totaltotals[1] = 'Totals';
    refcell.offset(groupoffset,0,1,tablewidth).setValues([totaltotals]);
    offset++;
  }
}


function _setBaseMatrix(theSpreadsheet) {
  var actualsheet, parametersheet, metarange, metaindexes;
  var yearsindex, accounttableindex, accounttableendindex;
  var startYear, endYear;
  var indexTableHeadings, accountCodes,referenceYearColumnIndex;
  var accountTitleDictionary;
  var titlerowanchorcell, accountcolanchorcell;
  // get and clear actual data sheet
  actualsheet = theSpreadsheet.getSheetByName('Consolidated Actual');
  theSpreadsheet.setActiveSheet(actualsheet);
  actualsheet.clear();
//  actualsheet.getRange(1,1).setValue('hello actual');
  // collect various indexes from parameter sheet
  parametersheet = theSpreadsheet.getSheetByName('Parameters');
  metarange = parametersheet.getRange('A:A');
  metaindexes = metarange.getValues();
  metaindexes = _flattenValues(metaindexes);
  yearsindex = metaindexes.indexOf('Range')+1;
  // main indexes
  accounttableindex = metaindexes.indexOf('IndexTable')+1;
  accounttableendindex = metaindexes.indexOf('IndexTableEnd')+1;
  startYear = Math.round(parametersheet.getRange(yearsindex,2).getValue());
  endYear = Math.round(parametersheet.getRange(yearsindex,3).getValue());
  // collect table headings and reference year index from parameter sheet
  indexTableHeadings = parametersheet.getRange(accounttableindex + ":" + accounttableindex).getValues();
  indexTableHeadings = indexTableHeadings[0]; // get row values;
  referenceYearColumnIndex = indexTableHeadings.indexOf(endYear)+1;
  // collect reference year account codes from parameter sheet, and account titles from reference year sheet
  accountCodes = parametersheet.getRange(
    accounttableindex + 1, // row
    referenceYearColumnIndex, // col
    accounttableendindex - accounttableindex // length, not including table column headers
  ).getValues();
  accountCodes = _flattenValues(accountCodes);
  accountCodes = _compactArray(accountCodes);
  accountTitleDictionary = _getReferenceYearDictionary(theSpreadsheet,endYear,accountCodes);
  // generate matrix labels framework
  var actualtitlerowoffset = 1;
  var actualtitlecoloffset = 2;
  titlerowanchorcell = actualsheet.getRange(actualtitlerowoffset + 1,actualtitlecoloffset + 1);
  accountcolanchorcell = actualsheet.getRange(3,1);
  var count = (endYear - startYear) + 1;
  var year = startYear;
  // generate year column headers
  actualsheet.getRange(actualtitlerowoffset + 1,1).setValue('BaseData');
  for (var offset = 0; offset < count; offset++) {
    titlerowanchorcell.offset(0,offset).setValue(year);
    year++;
  }
  var offset = 0;
  // generate account and title row labels
  for (account in accountTitleDictionary) {
    accountcolanchorcell.offset(offset,0).setValue(account);
    accountcolanchorcell.offset(offset,1).setValue(accountTitleDictionary[account]);
    offset++;
  }
  // generate matrix numbers
  var valuefunnel = _getValueFunnel(parametersheet,indexTableHeadings,startYear,endYear,accounttableindex,accounttableendindex);
  var actualtitlerowindex = actualtitlerowoffset + 1;
  var actualtitleindexes = actualsheet.getRange(actualtitlerowindex + ":" + actualtitlerowindex).getValues();
  var actualtitleindexes = actualtitleindexes[0];
  var parameters = {
    // general parms
    spreadsheet: theSpreadsheet,
    parametersheet:parametersheet,
    actualsheet:actualsheet,
    actualtitlerowoffset:actualtitlerowoffset,
    actualtitlecoloffset:actualtitlecoloffset,
    actualtitleindexes:actualtitleindexes,
    refaccountcodes:accountCodes,
    valuefunnel:valuefunnel,
    endyear:endYear,
    // actual sheet parms
    actualsheet: actualsheet,
    titlerowanchorcell:titlerowanchorcell,
    accountcolanchorcell:accountcolanchorcell
  }
  // write out matrix numbers
  for (year = startYear; year <= endYear; year++) {
    parameters.year = year;
    _writeActualValues(parameters)
  }
}


function _getValueFunnel(parametersheet,indexTableHeadings,startYear,endYear,accounttableindex,accounttableendindex) {
  var valuefunnel = {};
  var accountcolcellcount = accounttableendindex - accounttableindex;
  for (var year = startYear; year <= endYear; year++) {
    var yearstring = year.toString();
    valuefunnel[yearstring] = {};
    var colref = indexTableHeadings.indexOf(year) + 1;
    var refcell = parametersheet.getRange(accounttableindex + 1,colref);
    for (var offset = 0; offset < accountcolcellcount; offset++) {
      var account = refcell.offset(offset,0).getValue();
      var rollinto = refcell.offset(offset,-1).getValue();
      if (account) {
        account = account.trim();
        var funnelentry;
        if (rollinto) {
          funnelentry = {};
          var rollaccount;
          rollinto = rollinto.split(';');
          for each (rollaccount in rollinto) {
            rollaccount = rollaccount.trim();
            rollaccount = rollaccount.split('=');
            rollaccount[1] = rollaccount[1] || 1;
            if (isNaN(rollaccount[1])) {
              rollaccount[1] = parseFloat(rollaccount[1]);
            }
            if (!isNaN(rollaccount[1])) {
              funnelentry[rollaccount[0].trim()] = rollaccount[1];
            }
          }
        } else {
          funnelentry = null;
        }
        valuefunnel[yearstring][account] = funnelentry;
      }
    }
  }
  return valuefunnel;
}


function _writeActualValues(parameters) {
  // collect values
  var theYear = parameters.year;
  var theEndYear = parameters.endyear;
  var theYearString = theYear.toString();
  var theSpreadsheet = parameters.spreadsheet;
  var theYearSheet = theSpreadsheet.getSheetByName(theYearString);
  var amountIndexMarkers = _flattenValues(theYearSheet.getRange('A:A').getValues());
  var accountIndexes = _flattenValues(theYearSheet.getRange('B:B').getValues());
  // remove metatag
  var index = accountIndexes.indexOf('AccountIndex');
  accountIndexes[index] = null;
  var theFunnel = parameters.valuefunnel;
  var values = {};
  var amountoffset = false;
  for (var offset in accountIndexes) {
    offset = parseInt(offset);
    var amounttag = amountIndexMarkers[offset];
    if (amounttag) { // amount col offset can change row to row
      var rowIndex = offset+1;
      var amountIndexRow = theYearSheet.getRange(rowIndex + ":" + rowIndex).getValues();
      amountIndexRow = amountIndexRow[0];
      amountoffset = amountIndexRow.indexOf('AMOUNT');
      if (amountoffset == -1) amountoffset = false;
    }
    var theAccount = accountIndexes[offset];
    if (theAccount) {
      if (amountoffset !== false) {
        var theValue = theYearSheet.getRange(offset + 1,amountoffset + 1).getValue();
        if (theValue) {
          var valueset = _getValueSet(theYear, theEndYear, theYearSheet, theAccount, theValue, theFunnel);
          for (var account in valueset) {
            if (values[account]) {
              values[account] += valueset[account];
            } else {
              values[account] = valueset[account];
            }
          }
        }
      }
    }
  }
  // write values
  var actualsheet = parameters.actualsheet;
  var actualtitlerowoffset = parameters.actualtitlerowoffset;
  var actualtitlecoloffset = parameters.actualtitlecoloffset;
  var actualtitleindexes = parameters.actualtitleindexes;
  var refaccountcodes = parameters.refaccountcodes;
  var refdatacol = actualtitleindexes.indexOf(theYear)+1;
  if (refdatacol == -1) return;
  var refdatarow = actualtitlerowoffset + 2;
  var refdatacell = actualsheet.getRange(refdatarow,refdatacol);
  for (account in values) {
    var cellrowoffset = refaccountcodes.indexOf(account);
    if (cellrowoffset != -1) {
      refdatacell.offset(cellrowoffset,0).setValue(values[account]);
    }
  }
}


// tranform the value through account transformations to confrom to referenceyear
function _getValueSet(theYear, theEndYear, theYearSheet, theAccount, theValue, theFunnel) {
  var valueset = {};
  var theNewValueset = {};
  var isValid = !isNaN(parseFloat(theValue)) && isFinite(theValue);
  if (!isValid) return valueset;
  valueset[theAccount] = theValue;
  for (var year = theYear; year < theEndYear; year++) { // do one transformation for each year prior to endyear
    var yearstring = year.toString();
    for (var account in valueset) {
      var theLookup = theFunnel[yearstring][account];
      if (theLookup) {
        var theValue = valueset[account];
        valueset[account] = null; // about to be replaced
        for (var theNewAccount in theLookup) {
          if (theLookup[theNewAccount]) {
            var theRatio = theLookup[theNewAccount];
            theNewValueset[theNewAccount] = theValue * theRatio;
          } else { // no transformation -- maintain old association
            theNewValueset[theNewAccount] = theValue;
          }
        }
        for (theNewAccount in theNewValueset) {
          if (valueset[theNewAccount]) {
            valueset[theNewAccount] += theNewValueset[theNewAccount];
          } else {
            valueset[theNewAccount] = theNewValueset[theNewAccount];
          }
        }
        theNewValueset = {};
      }
    }
  }
  for (var account in valueset) {
    if (!valueset[account]) delete valueset[account];
  }
  return valueset;
}


function _getReferenceYearDictionary(theSpreadsheet, endYear, accountCodes) {
  var referenceyearsheet = theSpreadsheet.getSheetByName(endYear.toString());
  var refAccountCodes;
  var accountTitleDictionary = {};
  var rowRef,title;
  refAccountCodes = referenceyearsheet.getRange('B:B').getValues(); // accountindex column
  refAccountCodes = _flattenValues(refAccountCodes);
  for each (account in accountCodes) {
    rowRef = refAccountCodes.indexOf(account)+1; // row index for account
    title = referenceyearsheet.getRange(rowRef, 3).getValue(); // one to the right
    accountTitleDictionary[account] = title;
  }
  return accountTitleDictionary;
}


function _getClassificationDictionary(theSpreadsheet, startMarker, endMarker) {
  var parametersheet = theSpreadsheet.getSheetByName('Parameters');
  var classificationDictionary = {};
  var markerColumn = parametersheet.getRange('A:A').getValues();
  markerColumn = _flattenValues(markerColumn);
  var startMarkerOffset = markerColumn.indexOf(startMarker);
  var endMarkerOffset = markerColumn.indexOf(endMarker);
  var classificationTable = parametersheet.getRange(startMarkerOffset+1,3,endMarkerOffset-startMarkerOffset+1,2).getValues();
  for each (var account in classificationTable) {
    classificationDictionary[account[0]] = account[1];
  }
  return classificationDictionary;
}
// flatten column values from contained single value arrays to array of values
function _flattenValues(list) {
  var flattened = [];
  for each(item in list) {
    flattened.push(item[0]);
  }
  return flattened;
}


function _compactArray(list) {
  var compacted = [];
  for each(item in list) {
    if (item) {
      compacted.push(item);
    }
  }
  return compacted;
}
// from http://www.willmaster.com/library/generators/currency-formatting.php
function _CurrencyFormat(number)
{
   var decimalplaces = 0;
   var decimalcharacter = "";
   var thousandseparater = ",";
   number = parseFloat(number);
   var sign = number < 0 ? "-" : "";
   var formatted = new String(number.toFixed(decimalplaces));
   if( decimalcharacter.length && decimalcharacter != "." ) { formatted = formatted.replace(/\./,decimalcharacter); }
   var integer = "";
   var fraction = "";
   var strnumber = new String(formatted);
   var dotpos = decimalcharacter.length ? strnumber.indexOf(decimalcharacter) : -1;
   if( dotpos > -1 )
   {
      if( dotpos ) { integer = strnumber.substr(0,dotpos); }
      fraction = strnumber.substr(dotpos+1);
   }
   else { integer = strnumber; }
   if( integer ) { integer = String(Math.abs(integer)); }
   while( fraction.length < decimalplaces ) { fraction += "0"; }
   temparray = new Array();
   while( integer.length > 3 )
   {
      temparray.unshift(integer.substr(-3));
      integer = integer.substr(0,integer.length-3);
   }
   temparray.unshift(integer);
   integer = temparray.join(thousandseparater);
   return sign + integer + decimalcharacter + fraction;
}
// based on http://ramblings.mcpher.com/Home/excelquirks/gooscript/transpose
function _transpose(values) {
  var transposed = Object.keys(values[0]).map ( function (columnNumber) {
      return values.map( function (row) {
        return row[columnNumber];
      });
    });
return transposed;
}