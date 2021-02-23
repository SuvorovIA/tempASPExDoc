$( document ).ready(function() {
    var dv = document.getElementById("divConPlace");
    var btn = document.createElement("button");
    btn.type = "button"
    btn.textContent = "testloadwordsimple";
    btn.onclick = function () {
        //$.ajax({
        //    type: "POST",
        //    contentType: "application/json; charset=utf-8",
        //    url: "Default.aspx/DownloadDoc",
        //    data: "{'id':'werwe'}",
        //    success: function (data) {


        //    },

        //    error: function () {

        //    },
        //    complete: function () {

        //    }
        //});
        $.ajax({
            type: "POST",
            url: "Default.aspx/DownloadFile",
            data: '{fileName: "sfdas.docx"}',
            contentType: "application/json; charset=utf-8",
            dataType: "json",
            success: function (r) {
                //Convert Base64 string to Byte Array.
                var bytes = Base64ToBytes(r.d);

                //Convert Byte Array to BLOB.
                var blob = new Blob([bytes], { type: "application/octetstream" });

                //Check the Browser type and download the File.
                var isIE = false || !!document.documentMode;
                if (isIE) {
                    window.navigator.msSaveBlob(blob, fileName);
                } else {
                    var url = window.URL || window.webkitURL;
                    link = url.createObjectURL(blob);
                    var a = $("<a />");
                    a.attr("download", "test.docx");
                    a.attr("href", link);
                    $("body").append(a);
                    a[0].click();
                    $("body").remove(a);
                }
            }
        });
    };

    dv.appendChild(btn);


    var btn2 = document.createElement("button");
    btn2.type = "button"
    btn2.textContent = "testloadExcelsimple";
    btn2.onclick = function () {
        $.ajax({
            type: "POST",
            url: "Default.aspx/DownloadFile",
            data: '{fileName: "sfdas.docx"}',
            contentType: "application/json; charset=utf-8",
            dataType: "json",
            success: function (r) {
                //Convert Base64 string to Byte Array.
                var bytes = Base64ToBytes(r.d);

                //Convert Byte Array to BLOB.
                var blob = new Blob([bytes], { type: "application/octetstream" });

                //Check the Browser type and download the File.
                var isIE = false || !!document.documentMode;
                if (isIE) {
                    window.navigator.msSaveBlob(blob, fileName);
                } else {
                    var url = window.URL || window.webkitURL;
                    link = url.createObjectURL(blob);
                    var a = $("<a />");
                    a.attr("download", "test.docx");
                    a.attr("href", link);
                    $("body").append(a);
                    a[0].click();
                    $("body").remove(a);
                }
            }
        });
    };
    dv.appendChild(btn2);



    var btn3 = document.createElement("button");
    btn3.type = "button"
    btn3.textContent = "testloadDocModified";
    btn3.onclick = function () {
        $.ajax({
            type: "POST",
            url: "Default.aspx/DownloadDoc",
            data: '{fileName: "sfdas.docx"}',
            contentType: "application/json; charset=utf-8",
            dataType: "json",
            success: function (r) {
                //Convert Base64 string to Byte Array.
                var bytes = Base64ToBytes(r.d);

                //Convert Byte Array to BLOB.
                var blob = new Blob([bytes], { type: "application/octetstream" });

                //Check the Browser type and download the File.
                var isIE = false || !!document.documentMode;
                if (isIE) {
                    window.navigator.msSaveBlob(blob, fileName);
                } else {
                    var url = window.URL || window.webkitURL;
                    link = url.createObjectURL(blob);
                    var a = $("<a />");
                    a.attr("download", "test.docx");
                    a.attr("href", link);
                    $("body").append(a);
                    a[0].click();
                    $("body").remove(a);
                }
            },
            error: function(){

            }
        });
    };
    dv.appendChild(btn3);

});



function DownloadFile(fileName) {
    $.ajax({
        type: "POST",
        url: "Default.aspx/DownloadFile",
        data: '{fileName: "' + fileName + '"}',
        contentType: "application/json; charset=utf-8",
        dataType: "json",
        success: function (r) {
            //Convert Base64 string to Byte Array.
            var bytes = Base64ToBytes(r.d);

            //Convert Byte Array to BLOB.
            var blob = new Blob([bytes], { type: "application/octetstream" });

            //Check the Browser type and download the File.
            var isIE = false || !!document.documentMode;
            if (isIE) {
                window.navigator.msSaveBlob(blob, fileName);
            } else {
                var url = window.URL || window.webkitURL;
                link = url.createObjectURL(blob);
                var a = $("<a />");
                a.attr("download", fileName);
                a.attr("href", link);
                $("body").append(a);
                a[0].click();
                $("body").remove(a);
            }
        }
    });
};
function Base64ToBytes(base64) {
    var s = window.atob(base64);
    var bytes = new Uint8Array(s.length);
    for (var i = 0; i < s.length; i++) {
        bytes[i] = s.charCodeAt(i);
    }
    return bytes;
};