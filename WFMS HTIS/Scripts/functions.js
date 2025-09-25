var redirectPathG = '';

function showError(errorMsg, time, redirectPath, clear) {
    errormsg(errorMsg, time);

    if (errorMsg.indexOf('Error') == -1 && errorMsg.indexOf('authorized') == -1) {
        redirectPathG = redirectPath;
    }
    else {
        redirectPathG = '';

    }
    $('#ProgressImg').hide();
}
function alert(errorMsg) {
    var time = 5;
    errormsg(errorMsg, time);
}
function errormsg(errorMsg, time) {
    if (errorMsg.indexOf('Error') >= -1) {
        $('.crossX').show();
        //$('.okX').show();
        $('#errorDiv').removeClass('alert-info');
        $('#errorDiv').removeClass('alert-danger');
        $('#errorDiv').removeClass('alert-success');
        $('#errorDiv').show().addClass('alert alert-danger');
    }
    else {
        $('#errorDiv').show().addClass('alert alert-success');
    }
    $('#errorDiv').html(errorMsg);

    setTimeout('hideError()', 3000);
    //setTimeout('hideError()', 1000 * parseInt(time));
}

function hideError() {
    $('.crossX').hide();
    $('.okX').hide();
    $('#errorDiv').html('').hide();
    //redirectTo(redirectPathG); 
}

function redirectTo(path) {

    if (path != '') {
        $('#errorDiv').html('Redirecting to List...').show();
        window.location = path;
    }
}
var btnShown = 'N';
var intrval;
//var i = 0;
var btn1;
function showLeaveTypeBtn() {
   // i = i + 1;
   // $('#grdleave').html(i);
    // alert(document.getElementById('roleId').value);
    //if (document.getElementById('roleId').value != null) {
    if (document.getElementById('roleId').value == 7) {
        if (document.getElementById(btn1) != null) {
            document.getElementById(btn1).style.display = 'none';
            btnShown = 'Y';
        }
        else {
            //$('#grdleave').html('x');
        } 
    }
    if (btnShown == 'Y') {
        clearInterval(intrval);
    }
    //}
}

function hideButtonInEmployee(btn) {
    //$('#' + btn).hide();
    btn1 = btn;
    intrval = setInterval('showLeaveTypeBtn()', 1000);
}

 


function filterGlobalClear() {
    $('.filterClear').val('');
    $('.modal-close').html('X').addClass('close');
}