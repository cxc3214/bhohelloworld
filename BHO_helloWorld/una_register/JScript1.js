
var sfz = {
      _toNum: function num(n) {
        return isNaN(parseInt(n)) ? 0 : parseInt(n)
    }

    , getsex: function sex(cardNO) {
        if (cardNO.length == 18) {
            return (cardNO.charAt(16) % 2 == 1) ? "1" : "2";
        }
        if (cardNO.length == 15) {
            return (cardNO.charAt(14) % 2 == 1) ? "1" : "2";
        }
    }
    , getBirthday: function getBirthdatByIdNo(iIdNo) {
        var tmpStr = "";
        var idDate = "";
        var tmpInt = 0;
        var strReturn = "";
        iIdNo = $.trim(iIdNo);
        if (iIdNo.length == 15) {
            tmpStr = iIdNo.substring(6, 12);
            tmpStr = "19" + tmpStr;
            tmpStr = tmpStr.substring(0, 4) + "-" + tmpStr.substring(4, 6) + "-" + tmpStr.substring(6)

            return tmpStr;
        }
        else if (iIdNo.length == 18) {
            tmpStr = iIdNo.substring(6, 14);
            tmpStr = tmpStr.substring(0, 4) + "-" + tmpStr.substring(4, 6) + "-" + tmpStr.substring(6)
            return tmpStr;
        }
    }
};

//alert(sfz.getsex("320682198404088271"));

$(function() {
    $("input[name=sfzbh]").bind("propertychange", function(e) {
        if (($(this).val().length == 15) || ($(this).val().length == 18)) {
            e.preventDefault();
            /*
            if (sex($(this).val()) == "1") {
            $("input[name=xb]").val('DBXB1');
            $("#_xb").val('男');
            }
            if (sex($(this).val()) == "2") {
            $("input[name=xb]").val('DBXB2');
            $("#_xb").val('女');
            }
            */
            $("input[name=csrq]").val(getBirthdatByIdNo($("input[name=sfzbh]").val()))
        }
    })
    $("input[name=gz],input[name=syf],input[name=ffyf],input[name=jj],input[name=zj],input[name=sybxj],input[name=yzbt],input[name=qt]").bind("propertychange", function(e) {
        $("input[name=ysr]").val(num($("input[name=gz]").val()) + num($("input[name=syf]").val()) + num($("input[name=ffyf]").val()) + num($("input[name=jj]").val()) + num($("input[name=zj]").val()) + num($("input[name=sybxj]").val()) + num($("input[name=yzbt]").val()) + num($("input[name=qt]").val()));
    })
});



