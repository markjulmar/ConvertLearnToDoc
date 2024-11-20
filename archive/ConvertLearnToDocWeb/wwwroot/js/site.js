// Please see documentation at https://docs.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

$(function () {
    $('.alert').alert();

    $(document).on('change', '.custom-file-input', function (event) {
        $(this).next('.custom-file-label').html(event.target.files[0].name);
    })

    $('a[data-toggle="tab"]').on('shown.bs.tab', function (_e) {
        $('.alert').hide();
    });

    $('button[type="submit"]').on('click', function (e) {
        $('.alert').hide();

        $('#progressModal').modal({
            keyboard: false,
            show: true,
            backdrop: 'static'
        })

        if (e.target.id == "convertLearnToDoc") {
            checkCookie('tdrid');
        }
        else {
            checkCookie('fdrid');
        }
    });

});

function checkCookie(id) {
    var cookieVal = getCookie($('input[name=' + id + ']').val());
    if (cookieVal == null) {
        setTimeout("checkCookie('" + id + "');", 1000);
    }
    else {
        $('#progressModal').modal('hide');
        $('input[name=' + id + ']').val(createUniqueId);
    }
}  

function getCookie(id) {
    var cookieArr = document.cookie.split(";");
    for (var i = 0; i < cookieArr.length; i++) {
        var cookiePair = cookieArr[i].split("=");
        if (id == cookiePair[0].trim()) {
            return decodeURIComponent(cookiePair[1]);
        }
    }
    return null;
}

function createUniqueId() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
        var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
        return v.toString(16);
    });
}