var selectedArr = [];
//To Display State List based on Selected Country//
function GeoFenceStateListDisplay() {
    $("#nonSelected").html("");
    if ($('#GeoFenceCountryId').val() == '0') {
        alert('Please Select Country !!');
        return;
    }
    $('#ProgressImg').show();
    $.post('/Admin/_GeoFenceSitesStateList', {
        'CountryId': $('#GeoFenceCountryId').val()
    }, function (data) {
        $('#GeoFenceStateId').html(data);
        $('#ProgressImg').hide();
    });
}
//To Display City List based on Selected State//
function GeoFenceCityListDisplay() {
    $("#nonSelected").html("");
    if ($('#GeoFenceStateId').val() == '0') {
        alert('Please Select State !!');
        return;
    }
    $('#ProgressImg').show();
    $.post('/Admin/_GeoFenceSitesCityList', {
        'StateId': $('#GeoFenceStateId').val()
    }, function (data) {
            $('#GeofenceCityId').html(data);
            $('#ProgressImg').hide();
    });
}

//To Display GeoFence Sites List according to Selected City//
function GeoFenceSitesListDisplay() {
    $("#nonSelected").html("");
    if ($('#GeoFenceStateId').val() == '0') {
        alert('Please Select State !!');
        return;
    }
    if ($('#GeofenceCityId').val() == '0') {
        alert('Please Select City !!');
        return;
    }
    if ($('#EmpId').val() == '0') {
        alert('Please Select Employee !!');
        return;
    }
    $('#ProgressImg').show();   
    $.post('/Admin/GeoFenceSitesListDisplay', {
        'GeofenceCityId': $('#GeofenceCityId').val(),
        'EmpId': $('#EmpId').val(),
    }, function (data) {
        var html = "";
        for (var i = 0; i < data.length; i++) {
            html = "";
            var found = false;
            $("#MappedSiteList").find(".rectangle1").each(function () {
                if ($(this).attr("id") == data[i].EmployeeId) {
                    found = true;
                }
            })
            if (!found) {
                html += '<div id="' + data[i].GeoFenceId + '" class="rectangle1" onclick="ShiftSite(this)">';
                html += '<div class="div_Emp""><span class="span_EmpImage"><img src="/images/pin_img.png" class="img-circle" alt="Image" style="width:20px" /></span>';
                html += '</div><div>';
                html += '<span class="span_EmpName_Code">' + data[i].GeoFenceLocation + '</span>';
                html += '<span class="span_EmpDesignation_Branch">' + data[i].GeoFenceAddress + '</span>';
                html += '<span class="span_EmpEmail">' + data[i].StateName + ', ' + data[i].CityName + '</span>';
                html += '<span class="span_EmpEmail">' + data[i].GeoFenceLatitude + ', ' + data[i].GeoFenceLongitude + '</span>';
                html += '</div></div>';
                $('#nonSelected').append(html);
            }
        }
        $('#ProgressImg').hide();   
    });
}

function ShiftSite(x) {
    if ($('#EmpId').val().trim() == 0) {
        alert('Please Select Employee !!');
        return;
    }
    if ($(x).parent().attr("id") == "nonSelected") {
        $(x).detach().appendTo('#selected');
        var id = $(x).attr("id");
        selectedArr.push(id);
    }
    else {
        $(x).detach().appendTo('#nonSelected');
        var id = $(x).attr("id");
        selectedArr = jQuery.grep(selectedArr, function (a) {
            return a !== id;
        });
    }
}

//To Save/Update GeoFence Mapping Sites-Employee//
function GeoFenceMappingSubmit() {
    var SelectedGeoFenceIds = selectedArr.join(",");
    if ($('#EmpId').val() == '0') {
        $('.modelalert').html('Please Select Employee !!');
        return;
    }
    if ($('#GeoFenceCountryId').val() == '0') {
        $('.modelalert').html('Please Select Country !!');
        return;
    }
    if ($('#GeoFenceStateId').val() == '0') {
        $('.modelalert').html('Please Select State !!');
        return;
    }
    if ($('#GeofenceCityId').val() == '0') {
        $('.modelalert').html('Please Select City !!');
        return;
    }
    if (SelectedGeoFenceIds.length == 0) {
        $('.modelalert').html('Please Select atleast one GeoFence Site !!');
        return;
    }
    $('.modelalert').html('');
    $('#ModalProgress').show();
    $('.btnModalSubmit').attr("disabled", "disabled");
    $.post('/Admin/GeoFenceMappingSaveUpdateSubmit',
        {
            'GeoFenceMappingId': $('#GeoFenceMappingId').val(),
            'EmpId': $('#EmpId').val(),
            'GeoFenceId': SelectedGeoFenceIds,
        }, function (data) {
            if (data.indexOf('Error') > -1) {
                $('.modelalert').html(data);
            }
            else {
                $('.modelalert').html(data).delay(1000).queue(function () {
                    $('.modelalert').html('closing...')
                    _gfmplist();
                    $('#myModal').modal('hide');
                    $('.modelalert').dequeue();
                });
            }
            $('.btnModalSubmit').removeAttr("disabled");
            $('#ModalProgress').hide();
        });
}