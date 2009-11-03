var dblc = {
    cz: "操作人员"
    , sp: "审批"
    , sh: "审核"

    , PSI: function PSI(url, t) {
        return $.ajax({
            type: "POST",
            url: url,
            data: "t=" + t + "",
            async: false
        }).responseText;
    }
    , request: function QueryString(fieldName) {
        var urlString = document.location.search;
        if (urlString != null) {
            var typeQu = fieldName + "=";
            var urlEnd = urlString.indexOf(typeQu);
            if (urlEnd != -1) {
                var paramsUrl = urlString.substring(urlEnd + typeQu.length);
                var isEnd = paramsUrl.indexOf('&');
                if (isEnd != -1) {
                    return paramsUrl.substring(0, isEnd);
                }
                else {
                    return paramsUrl;
                }
            }
            else {
                return null;
            }
        }
        else {
            return null;
        }
    }

}

$(function() {
    var flowId = dblc.request("flowId");

    if (flowId == "") {
        alert("流程ID失效");
        $("button[type=submit]").attr("disabled", true);
    } else {
        var reqlc = dblc.PSI("public/serverSideData/lzmz/dbzxInfo.jsp", flowId);
        var zt = $("input[name=lczt]").val().toLowerCase()
        alert(zt);
        if ($.trim(reqlc) == dblc.cz.toString()) {
            if (zt != "zt1") {
                alert("您没有权限操作此页面");
                $("button").attr("disabled", true);
                $("iframe").contents().find("button").attr("disabled", true);
            }
        }

        if ($.trim(reqlc) == dblc.sp.toString()) {
            if (zt != "zt1" && zt != "zt2") {
                alert("您没有权限操作此页面");
                $("button").attr("disabled", true);
                $("iframe").contents().find("button").attr("disabled", true);
            }
        }

        if ($.trim(reqlc) == dblc.sh.toString()) {
            if (zt != "zt2" && zt != "zt3") {
                alert("您没有权限操作此页面");
                $("button").attr("disabled", true);
                $("iframe").contents().find("button").attr("disabled", true);
            }
        }
    }
})