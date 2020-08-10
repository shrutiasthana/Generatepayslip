function consolidatedpdfspart() {
var uid = "======";
var ss = SpreadsheetApp.openById(uid);
var stt = ss.getSheetByName("CSV");
const datat = ss.getSheetByName("CSV").getDataRange().getValues();
const datam = ss.getSheetByName("Hybrid").getDataRange().getValues();
const datac = ss.getSheetByName("Coding").getDataRange().getValues();
const datai = ss.getSheetByName("Incentives").getDataRange().getValues();
const dataa = ss.getSheetByName("Adjustments").getDataRange().getValues();  

for (var i = 1; i <= datat.length; i++) {
    var tid = datat[i][0];
    var cellrange = stt.getRange(i+1, 13);

    //math info
    var header = ['Student_ID', 'Student_name', 'Program', 'Teacher_ID', 'Teacher_name', 'TDS', 'TDS_rate', 'Installments', 'Total'];
    var filtered = datam.filter(function (values) {
      return values[3] === tid;
      }); 
    var obj = filtered.map(function(values) {
      return header.reduce(function(o, k, i) {
        o[k] = values[i];
        return o; 
    }, {});
  });
  //coing info
  var header_c = ['Student_ID', 'Student_name', 'Program', 'Teacher_ID', 'Teacher_name', 'TDS', 'TDS_rate', 'Installments', 'Total'];
    var filtered_c = datac.filter(function (values_c) {
      return values_c[3] === tid;
      }); 
    var obj_c = filtered_c.map(function(values_c) {
      return header_c.reduce(function(o, k, i) {
        o[k] = values_c[i];
        return o; 
    }, {});
  });
  //incentive info
  var header_i = ['Teacher_id', 'Teacher_name', 'Description', 'TDS', 'TDS_rate', 'Net_Amount'];
    var filtered_i = datai.filter(function (values_i) {
      return values_i[0] === tid;
      }); 
    var obj_i = filtered_i.map(function(values_i) {
      return header_i.reduce(function(o, k, i) {
        o[k] = values_i[i];
        return o; 
    }, {});
  });
  //adjustment info
  var header_a = ['Teacher_id', 'Teacher_name', 'Description', 'TDS', 'TDS_rate', 'sign', 'Net_Amount'];
    var filtered_a = dataa.filter(function (values_a) {
      return values_a[0] === tid;
      }); 
    var obj_a = filtered_a.map(function(values_a) {
      return header_a.reduce(function(o, k, i) {
        o[k] = values_a[i];
        return o; 
    }, {});
  });
  //main info
    var header_t = ['user_id', 'payout_id_injune20', 'name', 'Bank_Name', 'Acc_Number', 'PAN', 'total_gross_injune20', 'sign', 'total_adjustment_injune20', 'total_incentive_injune20', 'Total_tds_injune20', 'Final_paid_injune20', 'PDF_link_injune20', 'email'];
    var filtered_t = datat.filter(function (values_t) {
      return values_t[0] === tid;
      }); 
    var obj_t = filtered_t.map(function(values_t) {
      return header_t.reduce(function(o, k, i) {
        o[k] = values_t[i];
        return o; 
    }, {});
  });

  insertData2('docid',obj, obj_c, obj_i, obj_a, obj_t, cellrange);
  
}
}
function findContainingTable2(element) {
  if (element == null){}
  else{
  if (element.getType() == 'TABLE') {
    return element;
  }
  var parent = element.getParent()
  if (parent) {
    return findContainingTable2(parent);
  }
}}

function findContainingTableRow2(element) {
 if (element == null){}
  else{
  if (element.getType() == 'TABLE_ROW') {
    return element;
  }
  var parent = element.getParent()
  if (parent) {
    return findContainingTableRow2(parent);
  }
}}

function findPlaceholder2(element, placeholder) {
if (element == null){}
  else{
  if (element.getNumChildren !== undefined) {
    for (var i=0;i<element.getNumChildren();i++) {
      var child = element.getChild(i);
    
      if (child.getType() == 'PARAGRAPH') {
     
        if (child.getText().indexOf(placeholder) > -1) {
          return child;
        }
      }
      var res = findPlaceholder2(child, placeholder);
      if (res) {
        return res;
      }
    }
  }
  return null;
}}

function to2decimal2(num) {
  return Math.round(num * 100) / 100;
}

function num2str2(value) {
   var decimals = 2;
   return value.toFixed(decimals);
}

function insertData2(documentId, obj, obj_c, obj_i, obj_a, obj_t, cellrange) {
  Logger.log(obj_t[0]);
  var newdocumentId = DriveApp.getFileById(documentId).makeCopy(obj_t[0].payout_id_injune20).getId();
  var targetDoc = DocumentApp.openById(newdocumentId);
  var body = DocumentApp.openById(newdocumentId).getBody();
  var footer = DocumentApp.openById(newdocumentId).getFooter();
  
  body.replaceText('#{ t_name }', obj_t[0].name);
  body.replaceText('#{ b_name }', obj_t[0].Bank_Name);
  body.replaceText('#{ a_number }', obj_t[0].Acc_Number);
  body.replaceText('#{ pan }', obj_t[0].PAN);
  body.replaceText('#{ gross }', obj_t[0].total_gross_injune20);
  body.replaceText('#{msign}', obj_t[0].sign);
  body.replaceText('#{ adjust }', obj_t[0].total_adjustment_injune20);
  body.replaceText('#{ incentive }', obj_t[0].total_incentive_injune20);
  body.replaceText('#{ tax }', obj_t[0].Total_tds_injune20);
  body.replaceText('#{ netpaid }', obj_t[0].Final_paid_injune20);
  footer.replaceText('#{pay_id}', obj_t[0].payout_id_injune20);

  
  var placeholder = findPlaceholder2(body, '#{ m_name1 }');
  var table = findContainingTable2(placeholder);
  var totalAmount = 0.0;
  var totalVAT = 0.0;
  var tableRow = findContainingTableRow2(placeholder);
    if (obj.length==0){ table.removeChild(tableRow); 
    table.replaceText("Total fee amount payable - ₹#{mtotal_tax_fee} tax applicable", "No records for the period.")
    table.replaceText("₹#{ mGross fee }", " ");}
    
  for (var i=obj.length;i>0;i--) {
    var tableRow = findContainingTableRow2(placeholder);
    
    if (i!=1) {
      tableRow = table.insertTableRow(obj.length-i+1, tableRow.copy());
    }
    var obj1 = obj[obj.length - i];
    tableRow.replaceText('#{ m_name1 }', obj1.Student_name);
    tableRow.replaceText('#{ mprog }', obj1.Program);
    tableRow.replaceText('#{mtax}', obj1.TDS);
    tableRow.replaceText('#{m_rate}', obj1.TDS_rate);
    tableRow.replaceText('#{ m_n }', obj1.Installments);
    tableRow.replaceText('#{ m_amount }', obj1.Total);

    totalAmount += parseFloat(obj1.Total);
    totalVAT += parseFloat(obj1.TDS);
    
  }
  body.replaceText('#{ mGross fee }', num2str2(to2decimal(totalAmount)));
  body.replaceText('#{mtotal_tax_fee}', num2str2(to2decimal(totalVAT)));
 
  var placeholder2 = findPlaceholder2(body, '#{ c_name1 }');
  var table2 = findContainingTable2(placeholder2);
  var totalAmount2 = 0.0;
  var totalVAT2 = 0.0;
  var tableRow2 = findContainingTableRow2(placeholder2);
    if (obj_c.length==0){ table2.removeChild(tableRow2); 
    table2.replaceText("Total fee amount payable - ₹#{ctotal_tax_fee} tax applicable", "No records for the period.")
    table2.replaceText("₹#{ cGross fee }", " ");}
    
  for (var i=obj_c.length;i>0;i--) {
    var tableRow2 = findContainingTableRow2(placeholder2);
    
    if (i!=1) {
      tableRow2 = table2.insertTableRow(obj_c.length-i+1, tableRow2.copy());
    }
    var obj_c2 = obj_c[obj_c.length - i];
    tableRow2.replaceText('#{ c_name1 }', obj_c2.Student_name);
    tableRow2.replaceText('#{ cprog }', obj_c2.Program);
    tableRow2.replaceText('#{ctax}', obj_c2.TDS);
    tableRow2.replaceText('#{c_rate}', obj_c2.TDS_rate);
    tableRow2.replaceText('#{ c_n }', obj_c2.Installments);
    tableRow2.replaceText('#{ c_amount }', obj_c2.Total);

    totalAmount2 += parseFloat(obj_c2.Total);
    totalVAT2 += parseFloat(obj_c2.TDS);
    
  }
  body.replaceText('#{ cGross fee }', num2str(to2decimal(totalAmount2)));
  body.replaceText('#{ctotal_tax_fee}', num2str(to2decimal(totalVAT2)));
  
  var placeholder3 = findPlaceholder2(body, '#{adj_description}');
  var table3 = findContainingTable2(placeholder3);
  var totalAmount3 = 0.0;
  var totalVAT3 = 0.0;
  var tableRow3 = findContainingTableRow2(placeholder3);
  if (obj_a.length==0){ table3.removeChild(tableRow3); 
    table3.replaceText("Total  Adjustments due to Transfers or Refunds - ₹#{atotal_tax_fee} tax applicable", "No records for the period.");
    table3.replaceText("₹#{ gross_adj }", " ");}
    
  for (var i=obj_a.length;i>0;i--) {
    var tableRow3 = findContainingTableRow2(placeholder3);
    
    if (i!=1) {
      tableRow3 = table3.insertTableRow(obj_a.length-i+1, tableRow3.copy());
    }
    var obj_a3 = obj_a[obj_a.length - i];
    tableRow3.replaceText('#{adj_description}', obj_a3.Description);
    tableRow3.replaceText('#{atax}', obj_a3.TDS);
    tableRow3.replaceText('#{c_rate}', obj_a3.TDS_rate);
    tableRow3.replaceText('#{asign}', obj_a3.sign);
    tableRow3.replaceText('#{ a_amount }', obj_a3.Net_Amount);
    
    totalAmount3 += parseFloat(obj_a3.Net_Amount);
    totalVAT3 += parseFloat(obj_a3.TDS);
    
  }
  body.replaceText('#{ gross_adj }', num2str2(to2decimal(totalAmount3)));
  body.replaceText('#{atotal_tax_fee}', num2str2(to2decimal(totalVAT3)));
  
  var placeholder4 = findPlaceholder2(body, '#{description}');
  var table4 = findContainingTable2(placeholder4);
  var totalAmount4 = 0.0;
  var totalVAT4 = 0.0;
  var tableRow4 = findContainingTableRow2(placeholder4);
    if (obj_i.length==0){ table4.removeChild(tableRow4); 
    table4.replaceText("Total incentive amount - ₹#{itotal_tax_inc} tax applicable", "No records for the period.")
    table4.replaceText("₹#{ gross_inc }", " ");}
    
  for (var i=obj_i.length;i>0;i--) {
    var tableRow4 = findContainingTableRow2(placeholder4);
    
    if (i!=1) {
      tableRow4 = table4.insertTableRow(obj_i.length-i+1, tableRow4.copy());
    }
    var obj_i4 = obj_i[obj_i.length - i];
    tableRow4.replaceText('#{description}', obj_i4.Description);
    //tableRow4.replaceText('#{st}', obj_i4.Student_name);
    tableRow4.replaceText('#{itax}', obj_i4.TDS);
    tableRow4.replaceText('#{i_rate}', obj_i4.TDS_rate);
    tableRow4.replaceText('#{ i_amount }', obj_i4.Net_Amount);
    
    totalAmount4 += parseFloat(obj_i4.Net_Amount);
    totalVAT4 += parseFloat(obj_i4.TDS);
    
  }
  body.replaceText('#{ gross_inc }', num2str2(to2decimal(totalAmount4)));
  body.replaceText('#{itotal_tax_inc}', num2str2(to2decimal(totalVAT4)));
  
  targetDoc.saveAndClose();
  var pdf = DriveApp.getFileById(newdocumentId).getAs("application/pdf");
  var folder = DriveApp.getFolderById('1j11Zvcwi8shA8W0ZyMyyr6mAVRA8cIUD');
  var id = folder.createFile(pdf).getUrl();
  cellrange.setValue(id);
}

