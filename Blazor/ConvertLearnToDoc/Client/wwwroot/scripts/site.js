async function saveFile(fileName, stream) {
    const arrayBuffer = await stream.arrayBuffer();
    const blob = new Blob([arrayBuffer]);
    const url = URL.createObjectURL(blob);
    const anchorElement = document.createElement("a");
    anchorElement.href = url;
    anchorElement.download = fileName ?? "";
    anchorElement.click();
    anchorElement.remove();
    URL.revokeObjectURL(url);
}

var workingDialog;

function showWorkingDialog(element) {
    $('.alert').hide();

    if (workingDialog === null || workingDialog === undefined) {
        workingDialog = new bootstrap.Modal(element,
            {
                keyboard: false,
                focus: true,
                backdrop: 'static'
            });
        workingDialog.show();
    }
}

function hideWorkingDialog() {

    if (workingDialog !== null && workingDialog !== undefined) {
        workingDialog.hide();
        //workingDialog.dispose();

        workingDialog = null;
    }
}

$(function () {
    $('.alert').alert();

    $(document).on('change', '.custom-file-input', function (event) {
        $(this).next('.custom-file-label').html(event.target.files[0].name);
    })

    $('a[data-toggle="tab"]').on('shown.bs.tab', function (_e) {
        $('.alert').hide();
    });
});