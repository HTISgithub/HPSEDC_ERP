function show(txtBox, valBox, IdBox, url, colId, colName, ColCode, func) {

    $("#" + valBox).hide();
    $("#" + txtBox).autocomplete({
        source: function (request, response) {
            $.ajax({
                url: url,
                data: JSON.stringify({
                    "txt": $('#' + txtBox).val(),
                    'type': $('#' + IdBox).val(),
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
            $("#" + txtBox).val(ui.item.label);
            $("#" + valBox).val(ui.item.value);
        },
        minLength: 1
    });
}