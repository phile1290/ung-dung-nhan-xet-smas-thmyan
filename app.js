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
        
        // Tối ưu hóa dung lượng gửi đi để đảm bảo ứng dụng chạy mượt, không bị lỗi giới hạn
        textTT27 = rawTT27.substring(0, 1500); 
        textThamChieu = rawThamChieu.substring(0, 1500);
        
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
    // SỬ DỤNG MODEL 'gemini-pro' ĐỂ ĐẢM BẢO ỔN ĐỊNH VÀ KHÔNG BỊ LỖI "NOT FOUND"
    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=${apiKey}`;

    for (let i = 0; i < danhSachHocSinh.length; i++) {
        let hs = danhSachHocSinh[i];
        
        document.getElementById('status').innerText = `Đang phân tích: ${hs.hoTen} (${i+1}/${danhSachHocSinh.length}).\nHệ thống đang xử lý cẩn thận, vui lòng không đóng trình duyệt...`;
        
        let promptText = "";
        if (loaiHienTai === "monHoc") {
            promptText = `Bạn là một chuyên gia đánh giá tiểu học chuyên nghiệp. 
            Mẫu tham chiếu văn phong: ${textThamChieu}
            Kết quả học tập các môn của học sinh: "${hs.diemSo}". 
            
            QUY TẮC BẤT DI BẤT DỊCH (BẮT BUỘC TUÂN THỦ 100%): 
            1. TUYỆT ĐỐI KHÔNG dùng đại từ: thầy, cô, giáo viên, học sinh, em, bạn, họ tên, người học, cá nhân, bản thân. 
            2. BẮT BUỘC DÙNG CÂU ẨN CHỦ NGỮ. (Ví dụ chuẩn: "Nắm vững kiến thức toán học, thực hành tính toán nhanh...").
            3. Nhận xét phải CHI TIẾT, BÁM SÁT VÀO ĐIỂM SỐ từng môn. Không viết chung chung. Mỗi học sinh là DUY NHẤT.
            4. Nhận xét thành ĐÚNG 3 CÂU SÚC TÍCH.
            5. CHỈ trả về đúng chuỗi JSON này, tuyệt đối không in thêm ký tự nào khác: {"nhanXet": "Câu 1.\\n\\nCâu 2.\\n\\nCâu 3."}`;
        } else {
            promptText = `Bạn là một chuyên gia đánh giá tiểu học chuyên nghiệp.
            Mẫu tham chiếu văn phong: ${textThamChieu}
            Kết quả học tập các môn của học sinh: "${hs.diemSo}".
            
            QUY TẮC BẤT DI BẤT DỊCH (BẮT BUỘC TUÂN THỦ 100%):
            1. TUYỆT ĐỐI KHÔNG dùng đại từ: thầy, cô, giáo viên, học sinh, em, bạn, họ tên, người học, cá nhân, bản thân.
            2. BẮT BUỘC DÙNG CÂU ẨN CHỦ NGỮ.
            3. Phân tích cụ thể dựa vào dữ liệu điểm.
            4. Viết thành 3 lĩnh vực (Năng lực chung, Năng lực đặc thù, Phẩm chất), mỗi lĩnh vực đúng 1 dòng ngắn gọn.
            5. CHỈ trả về đúng chuỗi JSON này, tuyệt đối không in thêm ký tự nào khác: {"nangLucChung": "...", "nangLucDacThu": "...", "phamChat": "..."}`;
        }

        let retries = 4; // Cấu hình tự động thử lại 4 lần nếu mạng chậm
        let success = false;
        let lastError = "";
        
        while (retries > 0 && !success) {
            try {
                const response = await fetch(apiUrl, {
                    method: "POST", headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ 
                        contents: [{ parts: [{ text: promptText }] }],
                        generationConfig: { temperature: 0.7 } 
                    })
                });
                
                if (!response.ok) {
                    const errDetail = await response.json();
                    throw new Error(errDetail.error.message || "Lỗi đường truyền API");
                }
                
                const resData = await response.json();
                let aiText = resData.candidates[0].content.parts[0].text;
                
                // Thuật toán ép kiểu JSON chuyên nghiệp (Tránh lỗi markdown của AI)
                let jsonMatch = aiText.match(/\{[\s\S]*\}/);
                if(!jsonMatch) throw new Error("Lỗi định dạng dữ liệu trả về");
                
                const ketQua = JSON.parse(jsonMatch[0]);
                
                if (loaiHienTai === "monHoc") {
                    hs.nhanXetMonHoc = ketQua.nhanXet;
                } else { 
                    hs.nangLucChung = ketQua.nangLucChung; hs.nangLucDacThu = ketQua.nangLucDacThu; hs.phamChat = ketQua.phamChat; 
                }
                success = true;
            } catch (err) {
                lastError = err.message;
                retries--;
                await delay(3500); // Tăng thời gian chờ lên 3.5 giây trước khi thử lại
            }
        }
        
        // Bắt lỗi triệt để, thông báo rõ ràng cho giáo viên
        if (!success) {
            console.error("Lỗi tại học sinh:", hs.hoTen, lastError);
            if (loaiHienTai === "monHoc") hs.nhanXetMonHoc = `[Hệ thống AI từ chối phản hồi - Vui lòng thử lại sau]\n\n...\n\n...`;
            else { hs.nangLucChung = `Lỗi hệ thống`; hs.nangLucDacThu = "..."; hs.phamChat = "..."; }
        }
        
        // Thời gian chờ 4.5 giây giữa mỗi học sinh để đảm bảo Google không chặn
        await delay(4500); 
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
    const doc = new docx.Document({ sections: [{ children: [tieuDeLop, table] }] });

    docx.Packer.toBlob(doc).then(blob => {
        const link = document.createElement("a");
        link.href = URL.createObjectURL(blob);
        link.download = tenTepXuat;
        link.click();
    });
}