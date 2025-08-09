document.querySelectorAll(".upload-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    const fileInput = btn.closest("label").querySelector(".file-input");
    fileInput.click();
  });
});

document.querySelectorAll(".file-input").forEach((input) => {
  input.addEventListener("change", () => {
    const fileNameDiv = input.closest("label").querySelector(".filename");
    if (input.files.length > 0) {
      fileNameDiv.textContent = input.files[0].name;
    } else {
      fileNameDiv.textContent = "Chưa có tệp";
    }
  });
});

// drag and drop
document.querySelectorAll(".box").forEach((box) => {
  const fileInput = box.closest("label").querySelector(".file-input");
  const fileNameDiv = box.querySelector(".filename");

  box.addEventListener("dragover", (e) => {
    e.preventDefault();
    box.classList.add("dragover");
  });

  box.addEventListener("dragleave", () => {
    box.classList.remove("dragover");
  });

  box.addEventListener("drop", (e) => {
    e.preventDefault();
    box.classList.remove("dragover");

    const files = e.dataTransfer.files;
    if (files.length > 0) {
      fileNameDiv.textContent = files[0].name;

      // Trick để đặt file vào input[type=file]
      const dataTransfer = new DataTransfer();
      dataTransfer.items.add(files[0]);
      fileInput.files = dataTransfer.files;
    }
  });
});

//---------function--------//

let generatedWb = null;

async function joinFiles() {
  const input = document.getElementById("fileND");

  const file = input.files[0];
  const arrayBuffer = await file.arrayBuffer();

  // Tạo giá trị ngày hôm qua
  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);
  const yesterdayDateString = yesterday.toLocaleDateString("vi-VN", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  });

  // Đọc file Nạp đầu
  const wbNap = XLSX.read(arrayBuffer, { type: "array" });
  const sheetNap = wbNap.Sheets[wbNap.SheetNames[0]];
  const dataNap = XLSX.utils.sheet_to_json(sheetNap, { header: 1 });

  // Lấy dữ liệu từ file Nạp đầu
  const colA = dataNap.slice(2).map((row) => [row[0] || ""]); // Lấy dữ liệu từ Từ A3 STT
  const colC = dataNap.slice(2).map((row) => [row[2] || ""]); // Từ C3 tên tk
  const colD = dataNap.slice(2).map((row) => [row[3] || ""]); // Từ D3 lấy họ tên
  const colG = dataNap.slice(2).map((row) => [row[6] || ""]); //NGT cột G
  const colI = dataNap.slice(2).map((row) => [row[8] || ""]); //đại lý cột I
  const colN = dataNap.slice(2).map((row) => [row[13] || ""]); //số lần nạp cột N
  const colO = dataNap.slice(2).map((row) => [row[14] || ""]); // tổng nạp cột O
  const colQ = dataNap.slice(2).map((row) => [row[16] || ""]); // tổng rút cột Q
  const colAI = dataNap.slice(2).map((row) => [row[34] || ""]); // lấy cột ngân hàng AH

  // Đã nạp lại hoặc ko
  const recharge = colN.map((item) => {
    if (item[0] > 1) {
      return (item[0] = "ĐÃ NẠP LẠI");
    } else {
      return (item[0] = "KHÔNG");
    }
  });

  // Định dạng chỉ lấy dữ liệu tên ngân hàng
  const colAIslice = colAI.map((item) => {
    if (typeof item[0] === "string") {
      return item[0].split(/\s|\n/)[0];
    }
    return item[0];
  });

  // Tạo workbook mới
  generatedWb = XLSX.utils.book_new();

  // Chuẩn bị dữ liệu result: cột C bắt đầu từ C2
  const wsData = [];
  wsData.push([
    "STT",
    "Ngày",
    "Tài khoản",
    "Họ tên",
    "Đại lý",
    "NGT",
    "Tổng nạp",
    "Tổng cược",
    "Tổng rút",
    "IP",
    "Ngân hàng",
    "Chi nhánh",
    "Khuyến mãi",
    "Thiết bị",
    "Link",
    "Sản phẩm",
    "Cách cược",
    "Nạp lại",
  ]);

  // Đưa toàn bộ colC vào cột C
  for (let i = 0; i < colC.length; i++) {
    wsData[i + 1] = [
      colA[i][0],
      yesterdayDateString,
      colC[i][0],
      colD[i][0],
      colI[i][0],
      colG[i][0],
      colO[i][0],
      "" /*tổng cược*/,
      colQ[i][0],
      "" /*IP*/,
      colAIslice[i],
      "" /*chi nhánh*/,
      "" /*khuyến mãi*/,
      "" /* thiết bị*/,
      "" /*link*/,
      "" /*sản phẩm*/,
      "" /*cách cược*/,
      recharge[i],
    ];
  }

  const ws = XLSX.utils.aoa_to_sheet(wsData);
  XLSX.utils.book_append_sheet(generatedWb, ws, "Sheet1");

  // Hiện nút tải xuống
  document.getElementById("downloadBtn").style.display = "inline-block";
}

document.getElementById("downloadBtn").addEventListener("click", () => {
  if (generatedWb) {
    XLSX.writeFile(generatedWb, "result.xlsx");
  } else {
    alert("Chưa có dữ liệu để tải!");
  }
});
