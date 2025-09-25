var selectedArr = [];
MappedEmpListDisplay();
function MappedEmpListDisplay() {
    $.post('/Admin/_CGMappedEmployees', {
        'CGId': $('#CGId').val()
    }, function (e) {
        $('#MappedEmpList').html(e);
        $("#MappedEmpList").find(".rectangle1").each(function () {
            selectedArr.push(this.id);
        });
    });
}
function EmpListDisplay() {
    $("#nonSelected").html("");
    if ($('#DepartmentId').val() == '0') {
        alert('Please Select Department !!');
        return;
    }
    $.post('/Admin/_EmpListDisplay', {
        'DepartmentId': $('#DepartmentId').val()
    }, function (e) {
        var a = "";
        for (var t = 0; t < e.length; t++) {
            a = "";
            var l = false;
            $("#MappedEmpList").find(".rectangle1").each(function () {
                if ($(this).attr("id") == e[t].EmployeeId) {
                    l = true;
                }
            })
            if (!l) {
                a += '<div id="' + e[t].EmployeeId + '" class="rectangle1" onclick="ShiftEmployee(this)">';
                a += '<div class="div_Emp""><span class="span_EmpImage"><img src="/EmpImg/' + e[t].EmployeeImage + '" class="img-circle" alt="Image" style="width:70px" /></span>';
                a += '</div><div>';
                a += '<span class="span_EmpName_Code">' + e[t].EmployeeName + ' (' + e[t].EmployeeCode + ')</span>';
                a += '<span class="span_EmpDesignation_Branch">' + e[t].EmployeeDesignation + ',' + '<b>' + e[t].EmployeeBranch + '</b>' + '</span>';
                a += '<span class="span_EmpEmail">' + e[t].EmployeeEmail + '</span>';
                a += '</div></div>';
                $('#nonSelected').append(a);
            }
        }
    });
}
function ShiftEmployee(e) {
    if ($('#DepartmentId').val().trim() == 0) {
        alert('Please Select Department !!');
        return;
    }
    if ($(e).parent().attr("id") == "nonSelected") {
        $(e).detach().appendTo('#selected');
        var id = $(e).attr("id");
        selectedArr.push(id);
    }
    else {
        $(e).detach().appendTo('#nonSelected');
        var id = $(e).attr("id");
        selectedArr = jQuery.grep(selectedArr, function (a) {
            return a !== id;
        });
    }
}
function ClearanceGroupMappingSubmit() {
    var e = selectedArr.join(",");
    if ($('#DepartmentId').val() == '0') {
        $('.modelalert').html('Please Select Department !!');
        return;
    }
    $('.modelalert').html('');
    $('#ModalProgress').show();
    $('.btnModalSubmit').attr("disabled", "disabled");
    $.post('/Admin/GroupandClearanceHeadMappingSaveUpdateSubmit',
        {
            'CGId': $('#CGId').val(),
            'Emps': e,
        }, function (e) {
            if (e.indexOf('Error') > -1) {
                $('.modelalert').html(e);
            }
            else {
                $('.modelalert').html(e).delay(1000).queue(function () {
                    $('.modelalert').html('closing...')
                    _cgmlist();
                    $('#myModal').modal('hide');
                    $('.modelalert').dequeue();
                });
            }
            $('.btnModalSubmit').removeAttr("disabled");
            $('#ModalProgress').hide();
        });
}