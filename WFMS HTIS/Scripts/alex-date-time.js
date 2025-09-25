
(function (jq) {

    /*    jq.fn.typeADate = function () {
        
        var thisob = $(this);
        thisob.wrap('<div style="display: inline-block; position:relative;">');
        thisob.after('<div style="position:absolute; left:0; right:0; top:0; bottom:0; cursor: pointer;" ></div>');
        
        thisob.val("DD/MM/YYYY");
        
        var reg = /(0[1-9]|[12][0-9]|3[01]|DD)[\/](0[1-9]|1[012]|MM)[\/](19[0-9][0-9]|20[0-9][0-9]|YYYY)/;
        
        thisob.next().on("click touchstart", function() {
            thisob.focus();
    
            if (reg.test(thisob.val()) === false) { thisob.val("DD/MM/YYYY"); }
            
            setTimeout( function() { thisob.get(0).setSelectionRange(0,1); }, 0);
            init = {keypress:0};
            thisob.unbind("keydown").keydown( function(e) {
                e.preventDefault();
                var thisk = e.keyCode || e.which;
                var thisval = thisob.val();
                if ((thisk >= 48 && thisk <= 57) || (thisk >= 96 && thisk <= 105))  {
                    e.preventDefault();
                    if (thisval.length == 10 && init.keypress < 10) {
                        if (thisk >= 48 && thisk <= 57) { var thiskey = thisk - 48; }
                        if (thisk >= 96 && thisk <= 105) { var thiskey = thisk - 96; }
                        init.keypress++;
                        var frontval = thisval.substr(0,init.keypress-1);
                        var backval = thisval.substr(init.keypress,thisval.length);
                        thisob.val(frontval+thiskey+backval);
                        if (init.keypress == 2) { init.keypress = 3; }
                        if (init.keypress == 5) { init.keypress = 6; }
                        setTimeout( function() { thisob.get(0).setSelectionRange(init.keypress,init.keypress); }, 30);
                    }
                    
                } else if (thisk == 8) {
                    e.preventDefault();
                    if (init.keypress> 0) {
                        if (init.keypress == 3 || init.keypress == 6) { init.keypress--; }
                        init.keypress--;
                        setTimeout( function() { thisob.get(0).setSelectionRange(init.keypress,init.keypress+1); }, 30);
                    }
    
                } else { 
                    e.preventDefault();
                }
            });
        });
        thisob.blur( function() {
            if (reg.test(thisob.val()) === false) { thisob.css({"border-color": "#E04427", "background-color": "#EDD7D3"}) }
        });
    }
    */
    jq.fn.typeATime = function () {

        var thisob = jq(this);
        thisob.wrap('<div style="display: inline-block; position:relative; width:100%">');
        thisob.after('<div style="position:absolute; left:0; right:0; top:0; bottom:0; cursor: pointer;" ></div>');

        //thisob.val("00:00");

        var reg = /(0[0-9]|1[0-9]|2[0-3]|HH)[:](0[0-9]|[1-5][0-9]|HH)/;
        var isLeftRight = 'N';
        var lastLeftRightKey='0';
        thisob.next().on("click touchstart", function () {
            thisob.focus();
            if (reg.test(thisob.val()) === false) { thisob.val("00:00"); }
            setTimeout(function () { thisob.get(0).setSelectionRange(0, 1); }, 10);
            init = { keypress: 0 };
            thisob.unbind("keydown").keydown(function (e) {
                var thisk = e.keyCode || e.which;
                var thisval = thisob.val();
                //$('#txtRemarks').val(thisk);
                if (thisk == 37) {
                    lastLeftRightKey =37;
                    isLeftRight = 'Y';
                    thisk = 8;
                    if (init.keypress == 5) {
                        init.keypress = 4;
                    }

                    if (init.keypress == 0) {
                        init.keypress = 5;
                    }
                }
                if (thisk == 39) {
                    isLeftRight = 'Y';
                    lastLeftRightKey = 39;
                    if (init.keypress == 0) {
                        init.keypress = 1;
                    }
                    if (init.keypress == 2) {
                        init.keypress = 3;
                    }
                    if (init.keypress == 5) {
                        init.keypress = 0;
                    }
                    e.preventDefault();
                    init.keypress++;
                    setTimeout(function () { thisob.get(0).setSelectionRange(init.keypress - 1, init.keypress); }, 10);
                    //$('#txtRemarks').val(thisk + '-' + init.keypress);
                    return;
                }
                
                if ((thisk >= 48 && thisk <= 57) || (thisk >= 96 && thisk <= 105)) {
                    e.preventDefault();
                    if (isLeftRight == 'Y' && lastLeftRightKey ==39)
                    {
                        init.keypress--;
                        isLeftRight = 'N';
                        lastLeftRightKey =0;
                    }

                    if (thisval.length == 5 && init.keypress < 5) {
                        if (thisk >= 48 && thisk <= 57) { var thiskey = thisk - 48; }
                        if (thisk >= 96 && thisk <= 105) { var thiskey = thisk - 96; }

                         
                        init.keypress++;
                        

                        var frontval = thisval.substr(0, init.keypress - 1);
                        var backval = thisval.substr(init.keypress, thisval.length);
                        thisob.val(frontval + thiskey + backval);
                        var tval = frontval + thiskey + backval;
                        var x = tval.split(':');
                        if (init.keypress == 2) {
                            init.keypress = 3;
                            if (parseInt(x[0]) > 23) {
                                alert('invalid time');
                                init.keypress--
                                init.keypress--
                                init.keypress--
                            }
                        }
                        if (init.keypress == 5) {


                            if (parseFloat(x[1]) > 59) {
                                alert('invalid time');
                                init.keypress--
                                init.keypress--
                            }
                        }
                        
                        setTimeout(function () { thisob.get(0).setSelectionRange(init.keypress, init.keypress); }, 30);
                    }
                } else if (thisk == 8) {
                    e.preventDefault();
                    if (init.keypress > 0) {
                        if (init.keypress == 3) { init.keypress--; }
                        init.keypress--;
                        setTimeout(function () { thisob.get(0).setSelectionRange(init.keypress, init.keypress + 1); }, 30);
                    }

                } else {
                    e.preventDefault();
                }
            }).keyup(function () { var thisval = thisob.val(); if (thisval.length > 5) { var newval = thisval.substr(0, 5); thisob.val(newval); } });
            thisob.blur(function () {
                if (reg.test(thisob.val()) === false) { thisob.css({ "border-color": "#E04427", "background-color": "#EDD7D3" }) }
                else
                {
                    thisob.css({ "border-color": "#ccc", "background-color": "#fff" })
                }
            });
        });
    }

}(jQuery));