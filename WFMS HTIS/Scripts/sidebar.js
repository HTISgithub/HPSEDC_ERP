var menuG='Admin';
function _sidebarJS(menu) {
    menuG = menu;
    $.post('/menus/_sidebar', {
        'Menu': menu
    }, function (data) {
        $('#sidebar').html(data); 
    });
} 
function setddlTxt() {
    $("#ddlRoleMenu option:contains(" + menuG + ")").attr('selected', 'selected');
    $('#MenuDiv' + menuG).show();
}
 