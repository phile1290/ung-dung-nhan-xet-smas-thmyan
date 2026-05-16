// Cấu hình worker cho thư viện PDF.js
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';

let danhSachHocSinh = [];
let textTT27 = "";
let textThamChieu = "";
let loaiHienTai = "";
let tenTepXuat = "";

// Hàm đọc text từ file PDF
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

// Hàm tạo thời gian chờ (delay) để tránh lỗi quá tải API của Google
const delay = ms => new Promise(res => setTimeout(res, ms));

async function batDauXuLy() {
    const apiKey = document.getElementById('apiKey').value.trim();
    const fileTT27 = document.getElementById('pdfTT27').files[0];
    const fileThamChieu = document.getElementById('pdfThamChieu').files[0];
    const filesExcel = document.getElementById('fileUpload').files;
    const selectLoai = document.getElementById('loaiBangDiem');
    
    loaiHienTai = selectLoai.value;
    const moTaLoai = selectLoai.options[selectLoai.selectedIndex].text;
    tenTepXuat = `NhanXet_${moTaLoai.replace(/ /g, "_")}.docx`;
    document.getElementById('tenTepHienThi').innerText = tenTepXuat;

    if (!apiKey || !fileTT27 || !fileThamChieu || filesExcel.length === 0) {
        return alert("Vui lòng nhập API Key và tải lên đầy đủ 3 loại tệp: PDF TT27, PDF Tham chiếu và Bảng điểm SMAS!");
    }

    try {
        document.getElementById('status').innerText = "Đang trích xuất nội dung từ các file PDF...";
        textTT27 = await extractTextFromPDF(fileTT27);
        textThamChieu = await extractTextFromPDF(fileThamChieu);
        
        document.getElementById('status').innerText = "Đang quét TẤT CẢ CÁC SHEET và tổng hợp điểm...";
        
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
                                cotHoTen = c;
                                dongBatDau = r + 1; 
                                break;
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
                                    if (!hocSinhTongHop[tenHS]) {
                                        hocSinhTongHop[tenHS] = [];
                                    }
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
        
        if(danhSachHocSinh.length === 0) {
            return alert("Không tìm thấy dữ liệu học sinh hợp lệ. Vui lòng kiểm tra lại file SMAS!");
        }

        await taoNhanXetVoiAI(apiKey);
        
    } catch (error) {
        console.error(error);
        alert("Có lỗi xảy ra trong quá trình đọc file. Vui lòng kiểm tra lại file Excel.");
    }
}

async function taoNhanXetVoiAI(apiKey) {
    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;

    for (let i = 0; i < danhSachHocSinh.length; i++) {
        let hs = danhSachHocSinh[i];
        
        // Thông báo tiến trình rõ ràng cho giáo viên biết ứng dụng vẫn đang chạy
        document.getElementById('status').innerText = `Hệ thống AI đang phân tích chi tiết: ${hs.hoTen} (${i+1}/${danhSachHocSinh.length}).\nVui lòng đợi, quá trình này cần thời gian để tránh quá tải...`;
        
        let promptText = "";
        if (loaiHienTai === "monHoc") {
            promptText = `Đóng vai chuyên gia đánh giá tiểu học theo Thông tư 27. 
            Nội dung Thông tư 27: ${textTT27.substring(0, 4000)}...
            Nhận xét tham chiếu: ${textThamChieu.substring(0, 3000)}...
            
            Đây là BẢNG ĐIỂM TỔNG HỢP NHIỀU MÔN của học sinh: "${hs.diemSo}". 
            
            Quy tắc BẤT DI BẤT DỊCH: 
            1. TUYỆT ĐỐI KHÔNG sử dụng các từ: thầy, cô, giáo viên, học sinh, em, bạn, họ tên, người học, cá nhân, bản thân. 
            2. PHẢI SỬ DỤNG CÂU ẨN CHỦ NGỮ. (Ví dụ đúng: "Tiếp thu bài nhanh, tính toán thành thạo...". Ví dụ sai: "Cá nhân tiếp thu bài nhanh...").
            3. Viết nhận xét bao quát các môn, rõ ràng, chi tiết nhưng SÚC TÍCH, NGẮN GỌN HƠN (mỗi ý khoảng 15-20 từ, không viết quá dài).
            4. Nhận xét thành ĐÚNG 3 CÂU. Giữa mỗi câu PHẢI CÓ một dòng trống (dùng ký tự \\n\\n).
            5. Mỗi học sinh phải có nhận xét DUY NHẤT, không trùng lặp.
            6. CHỈ trả về JSON định dạng: {"nhanXet": "Câu 1.\\n\\nCâu 2.\\n\\nCâu 3."}. Không thêm bất kỳ text nào khác.`;
        } else {
            promptText = `Đóng vai chuyên gia đánh giá tiểu học theo Thông tư 27.
            Nội dung TT27: ${textTT27.substring(0, 4000)}...
            Nhận xét tham chiếu: ${textThamChieu.substring(0, 3000)}...
            
            Kết quả tổng hợp các môn của học sinh: "${hs.diemSo}".
            
            Quy tắc BẤT DI BẤT DỊCH:
            1. TUYỆT ĐỐI KHÔNG sử dụng các từ: thầy, cô, giáo viên, học sinh, em, bạn, họ tên, người học, cá nhân, bản thân.
            2. PHẢI SỬ DỤNG CÂU ẨN CHỦ NGỮ (Ví dụ đúng: "Giao tiếp tự tin...").
            3. Viết súc tích, bao quát nhưng độ dài vừa phải. Không viết dài dòng.
            4. Viết thành 3 lĩnh vực (Năng lực chung, Năng lực đặc thù, Phẩm chất), mỗi lĩnh vực 1 dòng.
            5. CHỈ trả về JSON: {"nangLucChung": "...", "nangLucDacThu": "...", "phamChat": "..."}. Không thêm text nào.`;
        }

        let retries = 5; // Tăng số lần thử lại nếu mạng lỗi
        let success = false;
        
        while (retries > 0 && !success) {
            try {
                const response = await fetch(apiUrl, {
                    method: "POST", headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ 
                        contents: [{ parts: [{ text: promptText }] }],
                        generationConfig: { temperature: 0.7 } 
                    })
                });
                
                if (!response.ok) throw new Error("API Limit");
                
                const resData = await response.json();
                let aiText = resData.candidates[0].content.parts[0].text.replace(/```json/g, "").replace(/```/g, "").trim();
                const ketQua = JSON.parse(aiText);
                
                if (loaiHienTai === "monHoc") {
                    hs.nhanXetMonHoc = ketQua.nhanXet;
                } else { 
                    hs.nangLucChung = ketQua.nangLucChung; hs.nangLucDacThu = ketQua.nangLucDacThu; hs.phamChat = ketQua.phamChat; 
                }
                success = true;
            } catch (err) {
                retries--;
                await delay(2000); // Đợi 2 giây trước khi thử lại để tránh Google khóa chặn
                if (retries === 0) {
                    console.error("Lỗi dòng học sinh", hs.hoTen);
                    // Chỉ dùng khi thất bại hoàn toàn 5 lần liên tiếp
                    if (loaiHienTai === "monHoc") hs.nhanXetMonHoc = "Nắm vững kiến thức trọng tâm các môn học.\n\nKỹ năng thực hành và vận dụng đạt yêu cầu.\n\nHoàn thành tốt các nhiệm vụ được giao.";
                    else { hs.nangLucChung = "Có ý thức tự giác học tập"; hs.nangLucDacThu = "Vận dụng kiến thức linh hoạt"; hs.phamChat = "Hòa đồng với bạn bè"; }
                }
            }
        }
        // THỜI GIAN CHỜ QUAN TRỌNG: Đợi 4.5 giây giữa mỗi học sinh để Google không đánh dấu là spam/quá tải API
        await delay(4500); 
    }
    document.getElementById('status').innerText = "Hoàn tất phân tích và đánh giá toàn bộ dữ liệu!";
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
    const rows = [];
    if (loaiHienTai === 'monHoc') {
        rows.push(new docx.TableRow({ children: [
            new docx.TableCell({ children: [new docx.Paragraph({text: "Họ và tên", bold: true})] }),
            new docx.TableCell({ children: [new docx.Paragraph({text: "Nhận xét đánh giá môn học và HĐGD", bold: true})] })
        ]}));
        
        danhSachHocSinh.forEach(hs => {
            let paragraphs = hs.nhanXetMonHoc.split('\n').map(line => new docx.Paragraph(line));
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
                new docx.TableCell({ children: [new docx.Paragraph(hs.nangLucChung)] }),
                new docx.TableCell({ children: [new docx.Paragraph(hs.nangLucDacThu)] }),
                new docx.TableCell({ children: [new docx.Paragraph(hs.phamChat)] })
            ]}));
        });
    }
    
    const table = new docx.Table({ rows: rows, width: { size: 100, type: docx.WidthType.PERCENTAGE } });
    const doc = new docx.Document({ sections: [{ children: [table] }] });

    docx.Packer.toBlob(doc).then(blob => {
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = tenTepXuat;
        link.click();
    });
}