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

function showWorkingDialog() {
    const modalElement = document.getElementById("progressDialog");
    modalElement.classList.add("modalShow");
    const bsModal = new bootstrap.Modal(modalElement);

    modalElement.addEventListener('shown.bs.modal', () => {
        if (!modalElement.classList.contains("modalShow")) {
            // Work was finished before modal fade-in was completed.
            // Go ahead and hide it now that it's showing.
            bsModal.hide();
        }
    }, { once: true });

    bsModal.show();
}

function hideWorkingDialog() {
    const modalElement = document.getElementById("progressDialog");
    modalElement.classList.remove("modalShow");

    const bsModal = bootstrap.Modal.getInstance(modalElement);
    if (bsModal) {

        modalElement.addEventListener('hidden.bs.modal', () => {
            bsModal.dispose();
        }, { once: true });

        bsModal.hide();
    }
}