const fileInput = document.getElementById("file");
const uploadArea = document.getElementById("uploadArea");
const uploadBtn = document.getElementById("uploadBtn");
const fileInfo = document.getElementById("fileInfo");
const fileName = document.getElementById("fileName");
const fileSize = document.getElementById("fileSize");

let storedFile = null;

if (fileInput && uploadArea) {
// mengubah ukuran dalam byte
  function formatFileSize(bytes) {
    if (bytes === 0) return "0 Bytes";
    // basis konversi (1 KB = 1024 Bytes).
    const k = 1024;
    const sizes = ["Bytes", "KB", "MB", "GB"];
    // menentukan index unit yang tepat berdasarkan besar bytes ("Bytes", "KB", "MB", "GB").
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + " " + sizes[i];
  }

  // Drag and drop
  uploadArea.addEventListener("dragover", (e) => {
    e.preventDefault();
    uploadArea.classList.add("dragover");
  });

  uploadArea.addEventListener("dragleave", () => {
    uploadArea.classList.remove("dragover");
  });

  uploadArea.addEventListener("drop", (e) => {
    e.preventDefault();
    uploadArea.classList.remove("dragover");

    if (e.dataTransfer.files.length > 0) {
      fileInput.files = e.dataTransfer.files;
      fileInput.dispatchEvent(new Event("change"));
    }
  });

  fileInput.addEventListener("change", (e) => {
    if (e.target.files.length > 0) {
      const file = e.target.files[0];

      storedFile = file;

      // Update file info
      fileName.textContent = file.name;
      fileSize.textContent = formatFileSize(file.size);
      fileInfo.classList.add("show");
      uploadArea.classList.add("file-selected");

      // text format
      uploadArea.querySelector(".upload-format").style.marginTop = "0px";
      uploadArea.querySelector(".upload-format").style.marginBottom = "15px";

      // Animasi
      uploadArea.style.transform = "scale(1.02)";
      setTimeout(() => {
        uploadArea.style.transform = "scale(1)";
      }, 200);
    }
  });
}
// Download functions
// const nama = `{{ nama }}`;
// const jadwal = JSON.parse(`{{ jadwal | tojson | safe }}`);
// const patners = JSON.parse(`{{ patners | tojson | safe }}`);

// function downloadPNG() {
//   const jadwalTable = document.querySelector("#jadwalTable");
//   const patnerTable = document.querySelector("#patnerTable");

//   const container = document.createElement("div");
//   container.style.background = "white";
//   container.style.padding = "30px";
//   container.style.textAlign = "center";
//   container.style.display = "inline-block";
//   container.style.fontFamily =
//     "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif";

//   const title = document.createElement("h2");
//   title.textContent = `Jadwal Ngawas ${nama}`;
//   title.style.color = "#667eea";
//   title.style.marginBottom = "20px";
//   container.appendChild(title);

//   if (jadwalTable) {
//     const jadwalClone = jadwalTable.cloneNode(true);
//     container.appendChild(jadwalClone);
//   }

//   if (patnerTable) {
//     const partnerTitle = document.createElement("h3");
//     partnerTitle.textContent = "Daftar Partners";
//     partnerTitle.style.color = "#667eea";
//     partnerTitle.style.marginTop = "30px";
//     partnerTitle.style.marginBottom = "15px";
//     container.appendChild(partnerTitle);

//     const patnerClone = patnerTable.cloneNode(true);
//     container.appendChild(patnerClone);
//   }

//   document.body.appendChild(container);

//   html2canvas(container, {
//     backgroundColor: "#ffffff",
//     scale: 2,
//   }).then((canvas) => {
//     const link = document.createElement("a");
//     link.download = `Jadwal_${nama.replace(/\s+/g, "_")}.png`;
//     link.href = canvas.toDataURL("image/png");
//     link.click();

//     document.body.removeChild(container);
//   });
// }

// function downloadExcel() {
//   fetch("/download_excel", {
//     method: "POST",
//     headers: { "Content-Type": "application/x-www-form-urlencoded" },
//     body: new URLSearchParams({
//       jadwal: JSON.stringify(jadwal),
//       patners: JSON.stringify(patners),
//       nama: nama,
//     }),
//   })
//     .then((response) => response.blob())
//     .then((blob) => {
//       const link = document.createElement("a");
//       link.href = window.URL.createObjectURL(blob);
//       link.download = `Jadwal_${nama.replace(/\s+/g, "_")}.xlsx`;
//       link.click();
//     })
//     .catch((err) => alert("Gagal mengunduh Excel: " + err));
// }

// ambil data dari <script id="app-data">
const raw = document.getElementById("app-data")?.textContent || "{}";
let APP = {};
try {
  APP = JSON.parse(raw);
} catch (e) {
  console.warn("Gagal parse APP_DATA:", e);
  APP = {};
}

const nama = (APP.nama || "").toString();
const jadwal = Array.isArray(APP.jadwal) ? APP.jadwal : [];
const patners = Array.isArray(APP.patners) ? APP.patners : [];

// download PNG
function downloadPNG() {
  const jadwalTable = document.querySelector("#jadwalTable");
  const patnerTable = document.querySelector("#patnerTable");

  // bungkus sementara untuk dirender html2canvas
  const container = document.createElement("div");
  container.style.background = "#ffffff";
  container.style.padding = "30px";
  container.style.textAlign = "center";
  container.style.display = "inline-block";
  container.style.fontFamily =
    "'Segoe UI', Tahoma, Geneva, Verdana, sans-serif";

  // judul
  const title = document.createElement("h1");
  title.textContent = `Jadwal Mengawas - ${nama || ""}`.trim();
  title.style.color = "black";
  title.style.marginBottom = "12px";
  title.style.fontWeight = "bold";
  title.style.fontSize = "22px";
  container.appendChild(title);

  // clone tabel jadwal
  if (jadwalTable) {
    const jadwalClone = jadwalTable.cloneNode(true);
    container.appendChild(jadwalClone);
  }

  // clone tabel partners (jika ada)
  if (patnerTable) {
    const partnerTitle = document.createElement("h1");
    partnerTitle.textContent = "Daftar Partners";
    partnerTitle.style.color = "black";
    partnerTitle.style.marginTop = "26px";
    partnerTitle.style.marginBottom = "12px";
    partnerTitle.style.fontWeight = "bold";
    partnerTitle.style.fontSize = "22px";
    container.appendChild(partnerTitle);

    const patnerClone = patnerTable.cloneNode(true);
    container.appendChild(patnerClone);
  }

  document.body.appendChild(container);

  html2canvas(container, { backgroundColor: "#ffffff", scale: 2 })
    .then((canvas) => {
      const a = document.createElement("a");
      const safeName = (nama || "Asisten").replace(/\s+/g, "_");
      a.download = `Jadwal_${safeName}.png`;
      a.href = canvas.toDataURL("image/png");
      a.click();
      document.body.removeChild(container);
    })
    .catch((err) => {
      document.body.removeChild(container);
      alert("Gagal membuat PNG: " + err.message);
      console.error(err);
    });
}

// download Excel
function downloadExcel() {
  const body = new URLSearchParams({
    jadwal: JSON.stringify(jadwal),
    patners: JSON.stringify(patners),
    nama: nama || "",
  });

  fetch("/download_excel", {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  })
    .then((res) => {
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      return res.blob();
    })
    .then((blob) => {
      const a = document.createElement("a");
      const safeName = (nama || "Asisten").replace(/\s+/g, "_");
      a.href = URL.createObjectURL(blob);
      a.download = `Jadwal_${safeName}.xlsx`;
      a.click();
      URL.revokeObjectURL(a.href);
    })
    .catch((err) => {
      alert("Gagal mengunduh Excel: " + err.message);
      console.error(err);
    });
}

// ekspor ke global supaya bisa dipanggil dari onclick di HTML
window.downloadPNG = downloadPNG;
window.downloadExcel = downloadExcel;
