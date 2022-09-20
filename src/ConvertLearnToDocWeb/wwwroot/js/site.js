// Please see documentation at https://docs.microsoft.com/aspnet/core/client-side/bundling-and-minification
// for details on configuring this project to bundle and minify static web assets.

// Write your JavaScript code.
$(document).on('change', '.custom-file-input', function (event) {
    $(this).next('.custom-file-label').html(event.target.files[0].name);
})

$('.alert').alert();

$('a[data-toggle="tab"]').on('shown.bs.tab', function (e) {
    $('.alert').hide();
});

$('button[type="submit"]').on('click', function (e) {
    $('.alert').hide();

    $('#progressModal').modal({
        keyboard: false,
        show: true,
        backdrop: 'static'
    })
});