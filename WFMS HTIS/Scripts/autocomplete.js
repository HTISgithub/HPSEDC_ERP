function autoText(txtBox, valBox, url, roleName) {
    if ($('#' + txtBox).val().trim() == '')
    {
        $("#" + valBox).val('0');
    }
  $("#" + valBox).hide();
     $("#" + txtBox).autocomplete({
         source: function (request, response) { 
             $.ajax({
                 url: url,
                 data:JSON.stringify({
                     "txt": $('#' + txtBox).val(),
                     'roleName':roleName
                 }), 
                 dataType: "json",
                 type: "POST",
                 contentType: "application/json; charset=utf-8",
                 dataFilter: function (data) { return data; },
                 success: function (data) { 
                     response($.map(data, function (item) {
                         return { label: item.EmpName, value: item.EmpId };
                     }))
                 },
                 error: function (XMLHttpRequest, textStatus, errorThrown) {
                     alert(textStatus);
                 }
             });
         },
         autoFocus: true,
         select: function (event, ui) {
            event.preventDefault();
            $("#"+txtBox).val(ui.item.label);
            if (ui.item) {
                $("#" + txtBox).val(ui.item.label);
                $("#" + valBox).val(ui.item.value);
            }
            
         },
         focus: function(event, ui) { 
             event.preventDefault(); 
              
             if ($('#' + txtBox).val().trim() == '') {
                 alert('');
                 $("#" + valBox).val('0');
             }
             //$("#" + txtBox).val(ui.item.label);
             //$("#" + valBox).val(ui.item.value);
         },
         minLength: 2
     });
}

function autoTextReport(txtBox, valBox, url,colId,colName,ColCode,func) {

    $("#" + valBox).hide();
    $("#" + txtBox).autocomplete({
        source: function (request, response) {
            $.ajax({
                url: url,
                data: JSON.stringify({
                    "txt": $('#' + txtBox).val(),
                    'colId': colId,
                    'colName': colName,
                    'colCode': ColCode,
                    'function':func
                }),
                dataType: "json",
                type: "POST",
                contentType: "application/json; charset=utf-8",
                dataFilter: function (data) { return data; },
                success: function (data) {
                    response($.map(data, function (item) {
                        return { label: item.EmpName, value: item.EmpId };
                    }))
                },
                error: function (XMLHttpRequest, textStatus, errorThrown) {
                    alert(textStatus);
                }
            });
        },
        select: function (event, ui) {
            event.preventDefault();
            $("#" + txtBox).val(ui.item.label);
            if (ui.item) {
                $("#" + txtBox).val(ui.item.label);
                $("#" + valBox).val(ui.item.value);
            }
        },
        focus: function (event, ui) {
            event.preventDefault();
            $("#" + txtBox).val(ui.item.label);
            $("#" + valBox).val(ui.item.value);
        },
        minLength: 1
    });
}

function autoTextVoucher(txtBox, valBox,type, url, colId, colName, ColCode, func) {

    $("#" + valBox).hide();
    $("#" + txtBox).autocomplete({
        source: function (request, response) {
            $.ajax({
                url: url,
                data: JSON.stringify({
                    "txt": $('#' + txtBox).val(),
                    "Type": $('#' + type).val(),
                    'colId': colId,
                    'colName': colName,
                    'colCode': ColCode,
                    'function': func
                }),
                dataType: "json",
                type: "POST",
                contentType: "application/json; charset=utf-8",
                dataFilter: function (data) { return data; },
                success: function (data) {
                    response($.map(data, function (item) {
                        return { label: item.EmpName, value: item.EmpId };
                    }))
                },
                error: function (XMLHttpRequest, textStatus, errorThrown) {
                    alert(textStatus);
                }
            });
        },
        select: function (event, ui) {
            event.preventDefault();
            $("#" + txtBox).val(ui.item.label);
            
            if (ui.item) {
                $("#" + txtBox).val(ui.item.label);
                $("#" + valBox).val(ui.item.value);
            }
        },
        focus: function (event, ui) {
            event.preventDefault();
            //$("#" + txtBox).val(ui.item.label);
            //$("#" + valBox).val(ui.item.value);
        },
        minLength: 1
    });
}