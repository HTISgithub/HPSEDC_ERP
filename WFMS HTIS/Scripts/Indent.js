$(document).ready(function(){
    $("#btnLaodReport").click(function()
    {ReportManager.LoadReport();}
    );

});
var ReportManager={LoadReport:function()
{
    var jsonParam="";
    var serviceurl = "/Indent/GenerateIndentList/";
    ReportManager.GetReport(serviceurl,jsonParam, onFailed);
  
    function onFailed(error){
        alert("Found error");

    }
},
    GetReport:function(serviceurl,jsonParam,errorCallback)
    {
 
        jQuery.ajax({
            url: serviceurl,
            async:false,
            type: "Post",
            data:"{" +jsonParam+"}",
            contentType:"application/json;Charset=utf-8",
            success:function(){
                window.open('/Report/ReportViewer.aspx', '_newtab');
                
            },
            
                error:errorCallback
                       
        });

    }

}