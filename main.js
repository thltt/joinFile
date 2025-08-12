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

      const dataTransfer = new DataTransfer();
      dataTransfer.items.add(files[0]);
      fileInput.files = dataTransfer.files;
    }
  });
});

//---------function Ghép file--------//

let generatedWb = null;

async function joinFiles() {
  const loader = document.getElementById("loader");
  loader.style.visibility = "visible";

  // lấy file
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

  // tìm index cột dựa vào từ khóa chứa trong tiêu đề file nạp đầu
  const headerRow = dataNap[1];
  function findColumnIndex(keyword) {
    return headerRow.findIndex((title) => title && title.toLowerCase().includes(keyword.toLowerCase()));
  }
  function safeIndexNap(keyword, defaultIndex) {
    const idx = findColumnIndex(keyword);
    return idx >= 0 ? idx : defaultIndex;
  }

  const indexLanNap = safeIndexNap("lần nạp", 13);
  const indexTongNap = safeIndexNap("tiền nạp", 14);
  const indexTongRut = safeIndexNap("tiền rút", 16);
  const indexNganHang = safeIndexNap("ngân hàng", 36);
  const indexIP = safeIndexNap("ip đăng nhập cuối");
  const indexLink = safeIndexNap("tên miền đăng nhập lần cuối", 25);
  const indexNGT = safeIndexNap("người giới thiệu", 6);
  const indexDL = safeIndexNap("đại lý", 8);
  const indexName = safeIndexNap("họ tên", 3);
  const indexUsername = safeIndexNap("tên tài khoản", 2);
  // Lấy dữ liệu từ file Nạp đầu
  const colA = dataNap.slice(2).map((row) => [row[0] || ""]); // Lấy dữ liệu từ Từ A3 STT
  const colC = dataNap.slice(2).map((row) => [row[indexUsername] || ""]); // Từ C3 tên tk
  const colD = dataNap.slice(2).map((row) => [row[indexName] || ""]); // Từ D3 lấy họ tên
  const colG = dataNap.slice(2).map((row) => [row[indexNGT] || ""]); //NGT cột G
  const colI = dataNap.slice(2).map((row) => [row[indexDL + 1] || ""]); //đại lý cột I
  const colN = dataNap.slice(2).map((row) => [row[indexLanNap] || ""]); //số lần nạp
  const colO = dataNap.slice(2).map((row) => [row[indexTongNap] || ""]); //tổng tiền nạp
  const colQ = dataNap.slice(2).map((row) => [row[indexTongRut] || ""]); //tổng tiền rút
  const colAI = dataNap.slice(2).map((row) => [row[indexNganHang] || ""]); //ngân hàng
  const colNapIP = dataNap.slice(2).map((row) => [row[indexIP] || ""]); // lấy cột IP
  const colNapLink = dataNap.slice(2).map((row) => [row[indexLink] || ""]); // lấy cột link

  // Lấy dữ liệu từ file đăng nhập
  const colUsername = dataDN.slice(2).map((row) => [row[1] || ""]); //lấy tên đn từ cột A
  const colFP = dataDN.slice(2).map((row) => [row[5] || ""]); //lấy Link từ cột F

  // Lấy dữ liệu từ file Khuyến mãi
  const colKMUsername = dataKM.slice(2).map((row) => [row[2] || ""]); // lấy username từ cột C
  const colKMname = dataKM.slice(2).map((row) => [row[12] || ""]); // lấy tên KM từ cột M

  // tìm index cột dựa vào từ khóa chứa trong tiêu đề file BCTT
  const headerBCTT = dataBCTT[1];
  function findIndexBCTT(keyword) {
    return headerBCTT.findIndex((title) => title && title.toLowerCase().includes(keyword.toLowerCase()));
  }
  function safeIndexBCTT(keyword, defaultIndex) {
    const idx = findIndexBCTT(keyword);
    return idx >= 0 ? idx : defaultIndex;
  }

  const indexBCTTUsername = safeIndexBCTT("tên tài khoản", 4);
  const indexPlatform = safeIndexBCTT("nền tảng", 3);
  const indexCHL = safeIndexBCTT("cược hợp lệ", 18);
  // Lấy dữ liệu từ file BCTT
  const colBCTTUsername = dataBCTT.slice(2).map((row) => [row[indexBCTTUsername] || ""]); // lấy username từ cột E
  const colBCTTplatfrom = dataBCTT.slice(2).map((row) => [row[indexPlatform] || ""]); // lấy tên sảnh từ cột D
  const colBCTTchl = dataBCTT.slice(2).map((row) => [row[indexCHL] || ""]); // lấy tổng CHL từ cột S

  // so sánh dữ liệu từ file đăng nhập vs file BCTT để lấy  tổng CHL
  const totalCHLMap = new Map();
  const flatformMap = new Map();
  const platformChlSeen = new Set(); // lưu

  const filteredBCTT = [];

  for (let i = 0; i < colBCTTUsername.length; i++) {
    const username = colBCTTUsername[i][0];
    let platform = colBCTTplatfrom[i][0] || "";
    let chlRaw = colBCTTchl[i][0];

    // Chuẩn hóa
    platform = platform.replace(/[^A-Za-z0-9 ]/g, "").trim();
    const chl = Number(String(chlRaw).replace(/[^0-9.-]/g, "")) || 0;

    if (username && platform) {
      const keyCheck = `${username}__${platform}__${chl}`;

      if (!platformChlSeen.has(keyCheck)) {
        platformChlSeen.add(keyCheck);

        // Cộng dồn CHL
        const currentTotal = totalCHLMap.get(username) || 0;
        totalCHLMap.set(username, currentTotal + chl);

        // Lưu platform
        const allPlatform = flatformMap.get(username) || [];
        if (!allPlatform.includes(platform)) {
          allPlatform.push(platform);
        }
        flatformMap.set(username, allPlatform);

        // Lưu vào mảng dữ liệu sạch
        filteredBCTT.push({
          username,
          platform,
          chl,
        });
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

  //so sánh dữ liệu từ file đăng nhập với file nạp đầu để lấy IP,link,thiết bị
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
      colNapIP[i][0],
      colAIslice[i],
      "" /*chi nhánh*/,
      promotionName,
      fp,
      colNapLink[i][0],
      platform,
      "" /*cách cược*/,
      recharge[i],
    ];
  }

  const ws = XLSX.utils.aoa_to_sheet(wsData);
  XLSX.utils.book_append_sheet(generatedWb, ws, "Sheet1");

  // Hiện nút tải xuống
  document.getElementById("downloadBtn").style.display = "inline-block";

  loader.style.visibility = "hidden";
}

document.getElementById("downloadBtn").addEventListener("click", () => {
  if (generatedWb) {
    XLSX.writeFile(generatedWb, "file-ket-qua.xlsx");
  } else {
    alert("Chưa có dữ liệu để tải!");
  }
});
