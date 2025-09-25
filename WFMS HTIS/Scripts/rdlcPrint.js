$(document).ready(function () {
    $("#btnLaodReport").click(function () { ReportManager.LoadReport(); }
    );

});
var ReportManager = {
    LoadReport: function () {
        var jsonParam = "";
        var serviceurl = "/ReportRdlc/GenerateAttendanceList/";
        ReportManager.GetReport(serviceurl, jsonParam, onFailed);

        function onFailed(error) {
            alert("Found error");

        }
    },
    GetReport: function (serviceurl, jsonParam, errorCallback) {

        jQuery.ajax({
            url: serviceurl,
            async: false,
            type: "Post",
            data: "{" + jsonParam + "}",
            contentType: "application/json;Charset=utf-8",
            success: function () {
                window.open('../RDLCWebPage/RDLCPrint.aspx', '_newtab');

            },

            error: errorCallback

        });

    }

}