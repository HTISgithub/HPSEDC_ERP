

var handleDataTableDefault = function () {
    "use strict";

    if ($('#myTable').length !== 0) {
        $('#myTable').DataTable({
            dom: 'Bfrtip',
            buttons: [
                {
                    extend: 'colvis',
                    collectionLayout: 'fixed two-column'
                },
                {
                    extend: 'copyHtml5',
                    exportOptions: {
                        columns: ':visible'
                    }
                },
                {
                    extend: 'excelHtml5',
                    exportOptions: {
                        columns: ':visible'
                    }
                },
                {
                    extend: 'pdfHtml5',
                    exportOptions: {
                        columns: ':visible'
                    }
                }
            ],
            "columnDefs": [{
                "defaultContent": "-",
                "targets": "_all"
            }],
             
            responsive: true
        }).on('draw', function () {
            //FormSliderSwitcher.init();
        });
    }
};
var TableManageDefault = function () {
    "use strict";
    return {
        //main function
        init: function () {
            handleDataTableDefault();
        }
    };
}();



function switcStatus(root, id, status) {
    if (root.indexOf('/') < 0) {
        root = '/admin/' + root;
    }
    console.log('switcStatus');
    $.post(root,
        {
            'Id': id,
            'IsActive': status,
        },
        function (data) {
            alert(data);
        });
}



function showModal(root, title, large) {
    $('#myModalLayout').modal('show');
    if (large) {
        $('.modal-dialog').addClass('modal-lg');
    }

    $('.modal-title').html('<b>' + title + '</b>');
    $('#_view').html('');
    $('#_view').html('<img src="/content/ajax-loader.gif" class="loaderimg" />');
    $("#myModalLayout").one('shown.bs.modal', function () {
        $.post(root, {
        }, function (data) {
            $('#_view').html(data);
            $(".datepicker").datepicker({
                format: "dd-M-yyyy",
                autoclose: true
            });
        });
    });
}
function showModal1(root, title, large, func) {
    $('#myModal').modal('show');
    if (large) {
        $('.modal-dialog').addClass('modal-lg');
    }

    $('.modal-title').html('<b>' + title + '</b>');
    $('#_view').html('');
    $('#_view').html('<img src="/content/ajax-loader.gif" class="loaderimg" />');
    $("#myModal").one('shown.bs.modal', function () {
        func();
    });
}


function modalShowEvent(fun) {
    $('#myModal').on('shown.bs.modal', function () {
        console.log('asdf');
        fun();
    })
}

function showfooter() {
    $('.modal-dialog').css('width', '900px');
    $('.modal-footer').css('display', 'block');
    $('.modal-body').css('padding-bottom', '100px');
    $('.modal-footer').css('margin-top', '-50px');
}

//$(window).load(function () {
//    //showExport();
//});

function tableOptions() {
    $('#myTable').DataTable({
        dom: 'Bfrtip',
        buttons: [
            {
                extend: 'colvis',
                collectionLayout: 'fixed two-column'
            },
            {
                extend: 'copyHtml5',
                exportOptions: {
                    columns: ':visible'
                }
            },
            {
                extend: 'excelHtml5',
                exportOptions: {
                    columns: ':visible'
                }
            },
            {
                extend: 'pdfHtml5',
                exportOptions: {
                    columns: ':visible'
                }
            }
        ]

    });
} 
 
//$('#myModal').on('shown.bs.modal', function () {
//    console.log('asdf'); 
//})
//function modalClose() {
//    $('.btnModalSubmit').unbind('click');
//    $('.modal-footer').hide();
//    $('#ModalProgress').hide();
//    $('.modelalert').html('');
//    $('.modal-dialog').removeClass('modal-lg');
//    console.log('closinghide');
//}

$(".datepicker").datepicker({
    format: "dd-M-yyyy",
    autoclose: true
});

function showCal() {
    $(".datepicker").datepicker({
        format: "dd-M-yyyy",
        autoclose: true
    });
}