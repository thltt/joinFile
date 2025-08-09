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

//---------function Ghép file--------//

const progressBar = document.getElementById("progressBar");
progressBar.style.display = "block"; // Hiện progress

function showProgress() {
  progressBar.style.visibility = "visible";
}
function hideProgress() {
  progressBar.style.visibility = "hidden";
  progressBar.value = 0;
}

function updateProgress(value) {
  progressBar.value = value;
}

showProgress();
updateProgress(0);

let generatedWb = null;

async function joinFiles() {
  const inputND = document.getElementById("fileND");
  const inputDN = document.getElementById("fileDN");
  const inputKM = document.getElementById("fileKM");
  const inputBCTT = document.getElementById("fileBCTT");

  const fileND = inputND.files[0];
  const arrayBufferND = await fileND.arrayBuffer();

  const fileDN = inputDN.files[0];
  const arrayBufferDN = await fileDN.arrayBuffer();

  const fileKM = inputKM.files[0];
  const arrayBufferKM = await fileKM.arrayBuffer();

  const fileBCTT = inputBCTT.files[0];
  const arrayBufferBCTT = await fileBCTT.arrayBuffer();

  updateProgress(5);

  // Đọc file Nạp đầu
  const wbNap = XLSX.read(arrayBufferND, { type: "array" });
  const sheetNap = wbNap.Sheets[wbNap.SheetNames[0]];
  const dataNap = XLSX.utils.sheet_to_json(sheetNap, { header: 1 });

  // Đọc file Đăng nhập
  const wbDN = XLSX.read(arrayBufferDN, { type: "array" });
  const sheetDN = wbDN.Sheets[wbDN.SheetNames[0]];
  const dataDN = XLSX.utils.sheet_to_json(sheetDN, { header: 1 });

  // Đọc file Khuyến mãi
  const wbKM = XLSX.read(arrayBufferKM, { type: "array" });
  const sheetKM = wbKM.Sheets[wbKM.SheetNames[0]];
  const dataKM = XLSX.utils.sheet_to_json(sheetKM, { header: 1 });

  // Đọc file BCTT
  const wbBCTT = XLSX.read(arrayBufferBCTT, { type: "array" });
  const sheetBCTT = wbBCTT.Sheets[wbBCTT.SheetNames[0]];
  const dataBCTT = XLSX.utils.sheet_to_json(sheetBCTT, { header: 1 });

  updateProgress(10);

  // tìm index cột dựa vào từ khóa chứa trong tiêu đề
  const headerRow = dataNap[1];
  function findColumnIndex(keyword) {
    return headerRow.findIndex((title) => title && title.toLowerCase().includes(keyword.toLowerCase()));
  }
  const indexLanNap = findColumnIndex("lần nạp");
  const indexTongNap = findColumnIndex("tiền nạp");
  const indexTongRut = findColumnIndex("tiền rút");
  const indexNganHang = findColumnIndex("ngân hàng");

  // Lấy dữ liệu từ file Nạp đầu
  const colA = dataNap.slice(2).map((row) => [row[0] || ""]); // Lấy dữ liệu từ Từ A3 STT
  const colC = dataNap.slice(2).map((row) => [row[2] || ""]); // Từ C3 tên tk
  const colD = dataNap.slice(2).map((row) => [row[3] || ""]); // Từ D3 lấy họ tên
  const colG = dataNap.slice(2).map((row) => [row[6] || ""]); //NGT cột G
  const colI = dataNap.slice(2).map((row) => [row[8] || ""]); //đại lý cột I
  const colN = dataNap.slice(2).map((row) => [row[indexLanNap] || ""]); //số lần nạp
  const colO = dataNap.slice(2).map((row) => [row[indexTongNap] || ""]); //tổng tiền nạp
  const colQ = dataNap.slice(2).map((row) => [row[indexTongRut] || ""]); //tổng tiền rút
  const colAI = dataNap.slice(2).map((row) => [row[indexNganHang] || ""]); //ngân hàng

  // Lấy dữ liệu từ file đăng nhập
  const colUsername = dataDN.slice(2).map((row) => [row[1] || ""]); //lấy tên đn từ cột A
  const colIP = dataDN.slice(2).map((row) => [row[2] || ""]); //lấy IP từ cột C
  const colLink = dataDN.slice(2).map((row) => [row[3] || ""]); //lấy Link từ cột D
  const colFP = dataDN.slice(2).map((row) => [row[5] || ""]); //lấy Link từ cột F

  // Lấy dữ liệu từ file Khuyến mãi
  const colKMUsername = dataKM.slice(2).map((row) => [row[2] || ""]); // lấy username từ cột C
  const colKMname = dataKM.slice(2).map((row) => [row[12] || ""]); // lấy tên KM từ cột M

  // Lấy dữ liệu từ file BCTT
  const colBCTTUsername = dataBCTT.slice(2).map((row) => [row[4] || ""]); // lấy username từ cột E
  const colBCTTplatfrom = dataBCTT.slice(2).map((row) => [row[3] || ""]); // lấy username từ cột D
  const colBCTTchl = dataBCTT.slice(2).map((row) => [row[18] || ""]); // lấy tên CHL từ cột S

  updateProgress(15);

  // so sánh dữ liệu từ file đăng nhập vs file khuyến mãi để lấy tên tổng CHL
  const totalCHLMap = new Map();
  for (let i = 0; i < colBCTTUsername.length; i++) {
    const username = colBCTTUsername[i][0];
    const chlRaw = colBCTTchl[i][0];
    const chl = Number(String(chlRaw).replace(/[^0-9.-]/g, "")) || 0;
    if (username) {
      const currentTotal = totalCHLMap.get(username) || 0;
      totalCHLMap.set(username, currentTotal + chl);
    }
  }

  // so sánh dữ liệu từ file đăng nhập vs file khuyến mãi để lấy tên tên sảnh
  const flatformMap = new Map();
  for (let i = 0; i < colBCTTUsername.length; i++) {
    const username = colBCTTUsername[i][0];
    let platform = colBCTTplatfrom[i][0];
    if (username && platform) {
      platform = platform.replace(/[^A-Za-z0-9 ]/g, "").trim();
      if (platform) {
        const allPlatform = flatformMap.get(username) || [];
        if (!allPlatform.includes(platform)) {
          allPlatform.push(platform);
        }
        flatformMap.set(username, allPlatform);
      }
    }
  }
  // ghép lại thành chuỗi
  for (const [username, platforms] of flatformMap.entries()) {
    flatformMap.set(username, platforms.join(", "));
  }

  // so sánh dữ liệu từ file đăng nhập vs file khuyến mãi để lấy tên KM
  const promotionMap = new Map();
  for (let i = 0; i < colKMUsername.length; i++) {
    const username = colKMUsername[i][0];
    let promoName = colKMname[i][0];
    if (promoName.includes("THƯỞNG NẠP ĐẦU")) promoName = "KMND";
    if (username) {
      promotionMap.set(username, promoName);
    }
  }

  updateProgress(20);

  //so sánh dữ liệu từ file đăng nhập với file nạp đầu để lấy IP,link,thiết bị
  const ipMap = new Map(); //IP
  for (let i = 0; i < colUsername.length; i++) {
    const username = colUsername[i][0];
    const ip = colIP[i][0];
    if (username && !ipMap.has(username)) {
      ipMap.set(username, ip);
    }
  }
  const linkMap = new Map(); //link
  for (let i = 0; i < colUsername.length; i++) {
    const username = colUsername[i][0];
    const link = colLink[i][0];
    if (username && !linkMap.has(username)) {
      linkMap.set(username, link);
    }
  }
  const fpMap = new Map(); //fp
  for (let i = 0; i < colUsername.length; i++) {
    const username = colUsername[i][0];
    const fp = colFP[i][0];
    if (username && !fpMap.has(username)) {
      fpMap.set(username, fp);
    }
  }

  // Đã nạp lại hoặc ko
  const recharge = colN.map((item) => {
    if (item[0] > 1) {
      return (item[0] = "ĐÃ NẠP LẠI");
    } else {
      return (item[0] = "KHÔNG");
    }
  });

  updateProgress(30);

  // Tạo giá trị ngày hôm qua
  const today = new Date();
  const yesterday = new Date(today);
  yesterday.setDate(today.getDate() - 1);
  const yesterdayDateString = yesterday.toLocaleDateString("vi-VN", {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  });

  // Định dạng chỉ lấy dữ liệu tên ngân hàng
  const colAIslice = colAI.map((item) => {
    if (typeof item[0] === "string") {
      return item[0].split(/\s|\n/)[0];
    }
    return item[0];
  });

  updateProgress(50);

  // Tạo workbook mới
  generatedWb = XLSX.utils.book_new();

  // Chuẩn bị dữ liệu hàng tiêu đề
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

  // Đưa dữ liệu vào file
  for (let i = 0; i < colA.length; i++) {
    const usernameNap = colC[i][0];
    const ip = ipMap.get(usernameNap) || "";
    const link = linkMap.get(usernameNap) || "";
    const fp = fpMap.get(usernameNap) || "";
    const promotionName = promotionMap.get(usernameNap) || "";
    const chl = totalCHLMap.get(usernameNap) || 0;
    const platform = flatformMap.get(usernameNap) || "";
    wsData[i + 1] = [
      colA[i][0],
      yesterdayDateString,
      colC[i][0],
      colD[i][0],
      colI[i][0],
      colG[i][0],
      colO[i][0],
      chl,
      colQ[i][0],
      ip,
      colAIslice[i],
      "" /*chi nhánh*/,
      promotionName,
      fp,
      link,
      platform /*sản phẩm*/,
      "" /*cách cược*/,
      recharge[i],
    ];
  }

  updateProgress(90);

  const ws = XLSX.utils.aoa_to_sheet(wsData);
  XLSX.utils.book_append_sheet(generatedWb, ws, "Sheet1");

  updateProgress(100);

  // Hiện nút tải xuống
  document.getElementById("downloadBtn").style.display = "inline-block";
}

document.getElementById("downloadBtn").addEventListener("click", () => {
  if (generatedWb) {
    XLSX.writeFile(generatedWb, "file-ket-qua.xlsx");
  } else {
    alert("Chưa có dữ liệu để tải!");
  }
});

hideProgress();
progressBar.value = 0;
