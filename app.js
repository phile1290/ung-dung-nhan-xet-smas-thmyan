// Cấu hình worker cho thư viện PDF.js
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';

let danhSachHocSinh = [];
let textTT27 = "";
let textThamChieu = "";
let loaiHienTai = "";
let tenTepXuat = "";
let tenLopThucTe = ""; 

async function extractTextFromPDF(file) {
    const arrayBuffer = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({data: arrayBuffer}).promise;
    let text = "";
    for (let i = 1; i <= pdf.numPages; i++) {
        const page = await pdf.getPage(i);
        const content = await page.getTextContent();
        text += content.items.map(item => item.str).join(" ") + " ";
    }
    return text;
}

const delay = ms => new Promise(res => setTimeout(res, ms));

async function batDauXuLy() {
    const apiKey = document.getElementById('apiKey').value.trim();
    const fileTT27 = document.getElementById('pdfTT27').files[0];
    const fileThamChieu = document.getElementById('pdfThamChieu').files[0];
    const filesExcel = document.getElementById('fileUpload').files;
    const selectLoai = document.getElementById('loaiBangDiem');
    
    loaiHienTai = selectLoai.value;
    tenLopThucTe = document.getElementById('tenLop').value.trim(); 
    
    // Tên file xuất ra linh hoạt dựa trên mô tả lựa chọn của giáo viên
    const moTaLoai = selectLoai.options[selectLoai.selectedIndex].text;
    tenTepXuat = `NhanXet_${tenLopThucTe ? tenLopThucTe + "_" : ""}${moTaLoai.replace(/ /g, "_")}.docx`;
    document.getElementById('tenTepHienThi').innerText = tenTepXuat;

    if (!apiKey || !fileTT27 || !fileThamChieu || filesExcel.length === 0 || !tenLopThucTe) {
        return alert("Vui lòng điền đủ tất cả các trường (Tên lớp, API Key, và Tệp dữ liệu)!");
    }

    try {
        document.getElementById('status').innerText = "Đang trích xuất nội dung từ các file PDF...";
        let rawTT27 = await extractTextFromPDF(fileTT27);
        let rawThamChieu = await extractTextFromPDF(fileThamChieu);
        
        textTT27 = rawTT27.substring(0, 1500); 
        textThamChieu = rawThamChieu.substring(0, 1500);
        
        document.getElementById('status').innerText = "Đang quét các Sheet và lọc dữ liệu học sinh...";
        let hocSinhTongHop = {};
        
        for (let i = 0; i < filesExcel.length; i++) {
            const arrayBuffer = await filesExcel[i].arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, {type: 'array'});
            
            workbook.SheetNames.forEach(sheetName => {
                const sheet = workbook.Sheets[sheetName];
                const rawData = XLSX.utils.sheet_to_json(sheet, {header: 1});
                
                let cotHoTen = -1;
                let dongBatDau = -1;

                for (let r = 0; r < rawData.length; r++) {
                    if (rawData[r]) {
                        for (let c = 0; c < rawData[r].length; c++) {
                            if (typeof rawData[r][c] === 'string' && rawData[r][c].toLowerCase().includes("họ và tên")) {
                                cotHoTen = c; dongBatDau = r + 1; break;
                            }
                        }
                    }
                    if (cotHoTen !== -1) break;
                }

                if (cotHoTen !== -1) {
                    for (let r = dongBatDau; r < rawData.length; r++) {
                        let ten = rawData[r][cotHoTen];
                        if (ten && typeof ten === 'string' && isNaN(ten) && ten.trim().length > 4) {
                            let tenHS = ten.trim();
                            let tenLower = tenHS.toLowerCase();
                            
                            // Lọc bỏ các dữ liệu rác từ SMAS
                            if (!tenLower.includes("trường") && !tenLower.includes("hiệu trưởng") && 
                                !tenLower.includes("giáo viên") && !tenLower.includes("họ và tên") &&
                                !tenLower.includes("người lập")) {
                                
                                let diemSoCuaMon = [];
                                for (let c = cotHoTen + 1; c < rawData[r].length; c++) {
                                    if (rawData[r][c] !== undefined && rawData[r][c] !== "") {
                                        diemSoCuaMon.push(rawData[r][c]);
                                    }
                                }
                                
                                if (diemSoCuaMon.length > 0) {
                                    if (!hocSinhTongHop[tenHS]) hocSinhTongHop[tenHS] = [];
                                    hocSinhTongHop[tenHS].push(`${sheetName}: ${diemSoCuaMon.join(", ")}`);
                                }
                            }
                        }
                    }
                }
            });
        }
        
        danhSachHocSinh = Object.keys(hocSinhTongHop).map(ten => ({
            hoTen: ten,
            diemSo: hocSinhTongHop[ten].join(" | ")
        }));
        
        if(danhSachHocSinh.length === 0) return alert("Không tìm thấy dữ liệu học sinh!");

        await taoNhanXetVoiAI(apiKey);
        
    } catch (error) {
        console.error(error);
        alert("Có lỗi xảy ra trong quá trình đọc file. Vui lòng kiểm tra lại file Excel.");
    }
}

async function taoNhanXetVoiAI(apiKey) {
    // SỬ DỤNG CHUẨN GEMINI 1.5 FLASH (Model đỉnh nhất hiện tại cho văn bản)
    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;

    for (let i = 0; i < danhSachHocSinh.length; i++) {
        let hs = danhSachHocSinh[i];
        
        document.getElementById('status').innerText = `Đang phân tích: ${hs.hoTen} (${i+1}/${danhSachHocSinh.length}).\nỨng dụng đang điều tiết tốc độ thông minh để tránh lỗi...`;
        
        let promptText = "";
        let jsonSchema = {};

        // THIẾT LẬP LỆNH CHUẨN XÁC
        if (loaiHienTai === "monHoc") {
            promptText = `Bạn là chuyên gia đánh giá tiểu học theo Thông tư 27. 
            Mẫu tham chiếu: ${textThamChieu}
            Điểm các môn: "${hs.diemSo}". 
            
            QUY TẮC BẤT DI BẤT DỊCH: 
            1. TUYỆT ĐỐI KHÔNG dùng đại từ: thầy, cô, giáo viên, học sinh, em, bạn, họ tên, người học, cá nhân, bản thân. 
            2. CÂU ẨN CHỦ NGỮ. Nhận xét CHI TIẾT, BÁM SÁT ĐIỂM SỐ. Mỗi học sinh là DUY NHẤT.
            3. Viết ĐÚNG 3 CÂU SÚC TÍCH. Giữa mỗi câu có 2 lần xuống dòng (\\n\\n).`;
            
            // Ép Google hiểu cấu trúc mong muốn
            jsonSchema = {
                "type": "object",
                "properties": { "nhanXet": { "type": "string" } },
                "required": ["nhanXet"]
            };
        } else {
            promptText = `Bạn là chuyên gia đánh giá tiểu học theo Thông tư 27.
            Mẫu tham chiếu: ${textThamChieu}
            Đánh giá của học sinh: "${hs.diemSo}".
            
            QUY TẮC BẤT DI BẤT DỊCH:
            1. TUYỆT ĐỐI KHÔNG dùng đại từ: thầy, cô, giáo viên, học sinh, em, bạn, họ tên, người học, cá nhân, bản thân.
            2. CÂU ẨN CHỦ NGỮ. Phân tích cụ thể, không trùng lặp.
            3. Viết 3 lĩnh vực, mỗi lĩnh vực 1 dòng ngắn gọn.`;
            
            jsonSchema = {
                "type": "object",
                "properties": {
                    "nangLucChung": { "type": "string" },
                    "nangLucDacThu": { "type": "string" },
                    "phamChat": { "type": "string" }
                },
                "required": ["nangLucChung", "nangLucDacThu", "phamChat"]
            };
        }

        let retries = 5; // Cho phép thử lại tối đa 5 lần
        let success = false;
        let lastError = "";
        let waitTime = 4000; // Thời gian chờ linh hoạt (Exponential backoff)
        
        while (retries > 0 && !success) {
            try {
                const response = await fetch(apiUrl, {
                    method: "POST", 
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ 
                        contents: [{ parts: [{ text: promptText }] }],
                        generationConfig: { 
                            temperature: 0.7,
                            responseMimeType: "application/json" // VŨ KHÍ HẠNG NẶNG: Ép chuẩn JSON 100%
                        } 
                    })
                });
                
                if (!response.ok) {
                    const errDetail = await response.json();
                    throw new Error(errDetail.error.message || "Quá tải máy chủ");
                }
                
                const resData = await response.json();
                const aiText = resData.candidates[0].content.parts[0].text;
                
                // Trực tiếp phân tích JSON do đã được ép kiểu từ hệ thống
                const ketQua = JSON.parse(aiText);
                
                if (loaiHienTai === "monHoc") {
                    hs.nhanXetMonHoc = ketQua.nhanXet;
                } else { 
                    hs.nangLucChung = ketQua.nangLucChung; 
                    hs.nangLucDacThu = ketQua.nangLucDacThu; 
                    hs.phamChat = ketQua.phamChat; 
                }
                success = true;
            } catch (err) {
                lastError = err.message;
                retries--;
                console.warn(`Lỗi tạm thời tại ${hs.hoTen}. Thử lại sau ${waitTime/1000}s...`, lastError);
                await delay(waitTime);
                waitTime += 2000; // Tăng dần thời gian chờ nếu mạng tiếp tục nghẽn
            }
        }
        
        // Nếu đã thử 5 lần vẫn lỗi, thông báo để người dùng biết chứ không in câu dự phòng giống nhau
        if (!success) {
            console.error("Lỗi thất bại hoàn toàn tại:", hs.hoTen);
            if (loaiHienTai === "monHoc") hs.nhanXetMonHoc = `[Lỗi hệ thống máy chủ Google - Vui lòng xử lý lại em này]`;
            else { hs.nangLucChung = `Lỗi hệ thống`; hs.nangLucDacThu = "..."; hs.phamChat = "..."; }
        }
        
        // Thời gian nghỉ an toàn tiêu chuẩn giữa các học sinh
        await delay(5000); 
    }
    
    document.getElementById('status').innerText = "✅ Hoàn tất phân tích và đánh giá toàn bộ dữ liệu!";
    hienThiBang();
}

function hienThiBang() {
    document.getElementById('previewArea').style.display = 'block';
    let html = '<table>';
    if (loaiHienTai === 'monHoc') {
        html += '<tr><th style="width: 25%;">Họ và tên</th><th>Nhận xét đánh giá môn học và HĐGD</th></tr>';
        danhSachHocSinh.forEach(hs => {
            html += `<tr><td style="font-weight: bold; color: #004085;">${hs.hoTen}</td><td style="white-space: pre-line;">${hs.nhanXetMonHoc}</td></tr>`;
        });
    } else {
        html += '<tr><th style="width: 20%;">Họ và tên</th><th>Đánh giá năng lực chung</th><th>Đánh giá năng lực đặc thù</th><th>Phẩm chất</th></tr>';
        danhSachHocSinh.forEach(hs => {
            html += `<tr><td style="font-weight: bold; color: #004085;">${hs.hoTen}</td><td>${hs.nangLucChung}</td><td>${hs.nangLucDacThu}</td><td>${hs.phamChat}</td></tr>`;
        });
    }
    html += '</table>';
    document.getElementById('tableContainer').innerHTML = html;
}

function xuatFileWord() {
    const tieuDeLop = new docx.Paragraph({
        alignment: docx.AlignmentType.CENTER,
        spacing: { after: 300 }, 
        children: [
            new docx.TextRun({
                text: `NHẬN XÉT TỰ ĐỘNG LỚP ${tenLopThucTe}`.toUpperCase(),
                bold: true,
                size: 32, // 16pt 
                font: "Times New Roman"
            })
        ]
    });

    const rows = [];
    if (loaiHienTai === 'monHoc') {
        rows.push(new docx.TableRow({ children: [
            new docx.TableCell({ children: [new docx.Paragraph({text: "Họ và tên", bold: true})] }),
            new docx.TableCell({ children: [new docx.Paragraph({text: "Nhận xét đánh giá môn học và HĐGD", bold: true})] })
        ]}));
        
        danhSachHocSinh.forEach(hs => {
            // Đảm bảo xuống dòng rõ ràng khi xuất Word
            let paragraphs = hs.nhanXetMonHoc ? hs.nhanXetMonHoc.split('\n').map(line => new docx.Paragraph(line)) : [new docx.Paragraph("")];
            rows.push(new docx.TableRow({ children: [
                new docx.TableCell({ children: [new docx.Paragraph(hs.hoTen)] }),
                new docx.TableCell({ children: paragraphs })
            ]}));
        });
    } else {
        rows.push(new docx.TableRow({ children: [
            new docx.TableCell({ children: [new docx.Paragraph({text: "Họ và tên", bold: true})] }),
            new docx.TableCell({ children: [new docx.Paragraph({text: "Đánh giá năng lực chung", bold: true})] }),
            new docx.TableCell({ children: [new docx.Paragraph({text: "Đánh giá năng lực đặc thù", bold: true})] }),
            new docx.TableCell({ children: [new docx.Paragraph({text: "Phẩm chất", bold: true})] })
        ]}));
        
        danhSachHocSinh.forEach(hs => {
            rows.push(new docx.TableRow({ children: [
                new docx.TableCell({ children: [new docx.Paragraph(hs.hoTen)] }),
                new docx.TableCell({ children: [new docx.Paragraph(hs.nangLucChung || "")] }),
                new docx.TableCell({ children: [new docx.Paragraph(hs.nangLucDacThu || "")] }),
                new docx.TableCell({ children: [new docx.Paragraph(hs.phamChat || "")] })
            ]}));
        });
    }
    
    const table = new docx.Table({ rows: rows, width: { size: 100, type: docx.WidthType.PERCENTAGE } });
    const doc = new docx.Document({ sections: [{ children: [tieuDeLop, table] }] });

    docx.Packer.toBlob(doc).then(blob => {
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = tenTepXuat;
        link.click();
    });
}