var selectedArr = [];
MappedEmpListDisplay();
function MappedEmpListDisplay() {
    $.post('/Admin/_HelpDeskMappedEmployees', {
        'HelpdeskGroupId': $('#HelpdeskGroupId').val()
    }, function (e) {
        $('#MappedEmpList').html(e);
        $("#MappedEmpList").find(".rectangle1").each(function () {
            selectedArr.push(this.id);
        });
    });
}
function EmpListViaDepartmentDisplay() {
    $("#nonSelected").html("");
    $.post('/Admin/_EmpListDisplay', {
        'DepartmentId': $('#DepartmentId').val()
    }, function (e) {
        var a = "";
        for (var p = 0; p < e.length; p++) {
            a = "";
            var s = false;
            $("#MappedEmpList").find(".rectangle1").each(function () {
                if ($(this).attr("id") == e[p].EmployeeId) {
                    s = true;
                }
            })
            if (!s) {
                a += '<div id="' + e[p].EmployeeId + '" class="rectangle1" onclick="ShiftEmployee(this)">';
                a += '<div class="div_Emp""><span class="span_EmpImage"><img src="/EmpImg/' + e[p].EmployeeImage + '" class="img-circle" alt="Image" style="width:70px" /></span>';
                a += '</div><div>';
                a += '<span class="span_EmpName_Code">' + e[p].EmployeeName + ' (' + e[p].EmployeeCode + ')</span>';
                a += '<span class="span_EmpDesignation_Branch">' + e[p].EmployeeDesignation + ',' + '<b>' + e[p].EmployeeBranch + '</b>' + '</span>';
                a += '<span class="span_EmpEmail">' + e[p].EmployeeEmail + '</span>';
                a += '</div></div>';
                $('#nonSelected').append(a);
            }
        }
    });
}
function EmpListViaDesignationDisplay() {
    $("#nonSelected").html("");
    $.post('/Admin/_EmpListViaDesignationDisplay', {
        'DesignationId': $('#DesignationId').val()
    }, function (e) {
        var a = "";
        for (var p = 0; p < e.length; p++) {
            a = "";
            var s = false;
            $("#MappedEmpList").find(".rectangle1").each(function () {
                if ($(this).attr("id") == e[p].EmployeeId) {
                    s = true;
                }
            })
            if (!s) {
                a += '<div id="' + e[p].EmployeeId + '" class="rectangle1" onclick="ShiftEmployee(this)">';
                a += '<div class="div_Emp""><span class="span_EmpImage"><img src="/EmpImg/' + e[p].EmployeeImage + '" class="img-circle" alt="Image" style="width:70px" /></span>';
                a += '</div><div>';
                a += '<span class="span_EmpName_Code">' + e[p].EmployeeName + ' (' + e[p].EmployeeCode + ')</span>';
                a += '<span class="span_EmpDesignation_Branch">' + e[p].EmployeeDesignation + ',' + '<b>' + e[p].EmployeeBranch + '</b>' + '</span>';
                a += '<span class="span_EmpEmail">' + e[p].EmployeeEmail + '</span>';
                a += '</div></div>';
                $('#nonSelected').append(a);
            }
        }
    });
}
function EmpListViaBranchDisplay() {
    $("#nonSelected").html("");
    $.post('/Admin/_EmpListViaBranchDisplay', {
        'BranchId': $('#BranchId').val()
    }, function (e) {
        var a = "";
        for (var p = 0; p < e.length; p++) {
            a = "";
            var s = false;
            $("#MappedEmpList").find(".rectangle1").each(function () {
                if ($(this).attr("id") == e[p].EmployeeId) {
                    s = true;
                }
            })
            if (!s) {
                a += '<div id="' + e[p].EmployeeId + '" class="rectangle1" onclick="ShiftEmployee(this)">';
                a += '<div class="div_Emp""><span class="span_EmpImage"><img src="/EmpImg/' + e[p].EmployeeImage + '" class="img-circle" alt="Image" style="width:70px" /></span>';
                a += '</div><div>';
                a += '<span class="span_EmpName_Code">' + e[p].EmployeeName + ' (' + e[p].EmployeeCode + ')</span>';
                a += '<span class="span_EmpDesignation_Branch">' + e[p].EmployeeDesignation + ',' + '<b>' + e[p].EmployeeBranch + '</b>' + '</span>';
                a += '<span class="span_EmpEmail">' + e[p].EmployeeEmail + '</span>';
                a += '</div></div>';
                $('#nonSelected').append(a);
            }
        }
    });
}
function ShiftEmployee(e) {
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
function HelpdeskGroupMappingSubmit() {
    var e = selectedArr.join(",");
    if (selectedArr.length == 0) {
        $('.modelalert').html('Please Select atleast one Employee !!');
        return;
    }
    $('.modelalert').html('');
    $('#ModalProgress').show();
    $('.btnModalSubmit').attr("disabled", "disabled");
    $.post('/Admin/HelpdeskMappingSaveUpdateSubmit',
        {
            'HelpdeskGroupId': $('#HelpdeskGroupId').val(),
            'Emps': e,
        }, function (e) {
            if (e.indexOf('Error') > -1) {
                $('.modelalert').html(e);
            }
            else {
                $('.modelalert').html(e).delay(1000).queue(function () {
                    $('.modelalert').html('closing...')
                    _HelpDeskGroupList();
                    $('#myModal').modal('hide');
                    $('.modelalert').dequeue();
                });
            }
            $('.btnModalSubmit').removeAttr("disabled");
            $('#ModalProgress').hide();
        });
}