// JS file
const form = document.getElementById("uploadForm");

if (form) {
    const loading = document.getElementById("loading");

    form.addEventListener("submit", async (e) => {
        e.preventDefault();

        loading.classList.remove("d-none");

        const formData = new FormData(form);

        const res = await fetch("/hoat-dong-khac", {
            method: "POST",
            body: formData
        });

        const blob = await res.blob();

        const url = window.URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = "FIT_HoatDongKhac.docx";
        a.click();

        loading.classList.add("d-none");
    });
}