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
    try {
        const apiKey = document.getElementById('apiKey').value.trim();
        const fileTT27 = document.getElementById('pdfTT27').files[0];
        const fileThamChieu = document.getElementById('pdfThamChieu').files[0];
        const filesExcel = document.getElementById('fileUpload').files;
        const selectLoai = document.getElementById('loaiBangDiem');
        
        loaiHienTai = selectLoai.value;
        const tenLopInput = document.getElementById('tenLop');
        tenLopThucTe = tenLopInput ? tenLopInput.value.trim() : "";
        
        const moTaLoai = selectLoai.options[selectLoai.selectedIndex].text;
        tenTepXuat = `NhanXet_${tenLopThucTe ? tenLopThucTe + "_" : ""}${moTaLoai.replace(/ /g, "_")}.docx`;
        document.getElementById('tenTepHienThi').innerText = tenTepXuat;

        if (!apiKey || !fileTT27 || !fileThamChieu || filesExcel.length === 0) {
            return alert("Vui lòng nhập API Key và tải lên đầy đủ các tệp dữ liệu!");
        }

        document.getElementById('status').style.color = "#0066cc";
        document.getElementById('status').innerText = "Đang trích xuất và tối ưu hóa nội dung từ các file PDF...";
        let rawTT27 = await extractTextFromPDF(fileTT27);
        let rawThamChieu = await extractTextFromPDF(fileThamChieu);
        
        textTT27 = rawTT27.substring(0, 3000); 
        textThamChieu = rawThamChieu.substring(0, 2000);
        
        document.getElementById('status').innerText = "Đang quét TẤT CẢ CÁC SHEET và tổng hợp dữ liệu...";
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
        
        if(danhSachHocSinh.length === 0) {
            return alert("Không tìm thấy dữ liệu học sinh hợp lệ. Vui lòng kiểm tra lại file SMAS!");
        }

        await taoNhanXetVoiAI(apiKey);
        
    } catch (error) {
        console.error(error);
        alert("Có lỗi xảy ra trong quá trình đọc file. Vui lòng kiểm tra lại file Excel.");
    }
}

// Hàm khởi tạo khung bảng hiển thị (Chạy ngay trước khi phân tích)
function khoiTaoBang() {
    document.getElementById('previewArea').style.display = 'block';
    const container = document.getElementById('tableContainer');
    
    let html = '<table id="bangKetQua">';
    if (loaiHienTai === 'monHoc') {
        html += '<thead><tr><th style="width: 25%;">Họ và tên</th><th>Nhận xét đánh giá môn học và HĐGD</th></tr></thead>';
    } else {
        html += '<thead><tr><th style="width: 20%;">Họ và tên</th><th>Đánh giá năng lực chung</th><th>Đánh giá năng lực đặc thù</th><th>Phẩm chất</th></tr></thead>';
    }
    html += '<tbody id="noiDungBang"></tbody></table>';
    container.innerHTML = html;
}

// Hàm thêm kết quả từng học sinh trực tiếp vào bảng (Hiển thị thời gian thực)
function themKetQuaLenBang(hs) {
    const tbody = document.getElementById('noiDungBang');
    let tr = document.createElement('tr');
    
    if (loaiHienTai === 'monHoc') {
        tr.innerHTML = `<td style="font-weight: bold; color: #004085;">${hs.hoTen}</td><td style="white-space: pre-line;">${hs.nhanXetMonHoc}</td>`;
    } else {
        tr.innerHTML = `<td style="font-weight: bold; color: #004085;">${hs.hoTen}</td><td>${hs.nangLucChung}</td><td>${hs.nangLucDacThu}</td><td>${hs.phamChat}</td>`;
    }
    tbody.appendChild(tr);
}

async function taoNhanXetVoiAI(apiKey) {
    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
    
    // Khởi tạo bảng ngay lập tức để giáo viên thấy danh sách bắt đầu chạy
    khoiTaoBang();

    // XỬ LÝ TUẦN TỰ TỪNG HỌC SINH (KHÔNG GỘP NHÓM)
    for (let i = 0; i < danhSachHocSinh.length; i++) {
        let hs = danhSachHocSinh[i];
        
        // Cập nhật trạng thái hiển thị rõ ràng, xuống dòng cho "Vui lòng đợi..."
        document.getElementById('status').style.color = "#0066cc";
        document.getElementById('status').innerText = `Ứng dụng đang xử lý chuyên sâu: ${hs.hoTen} (${i+1}/${danhSachHocSinh.length})\nVui lòng đợi...`;
        
        let promptText = "";
        let jsonSchema = {};

        if (loaiHienTai === "monHoc") {
            promptText = `Đóng vai chuyên gia đánh giá tiểu học theo Thông tư 27. 
            Mẫu tham chiếu: ${textThamChieu}
            
            BẢNG ĐIỂM TỔNG HỢP CỦA HỌC SINH NÀY:
            [Họ và tên: ${hs.hoTen}] - Kết quả: ${hs.diemSo}
            
            QUY TẮC BẤT DI BẤT DỊCH: 
            1. TUYỆT ĐỐI KHÔNG ghi các con số điểm (như 7,8,9,10) vào nội dung nhận xét.
            2. TUYỆT ĐỐI KHÔNG sử dụng các từ: thầy, cô, giáo viên, học sinh, em, bạn, họ tên, cá nhân, bản thân, người học. Bắt buộc dùng CÂU ẨN CHỦ NGỮ.
            3. Viết NGẮN GỌN, SÚC TÍCH, đi thẳng vào vấn đề kỹ năng thực tế.
            4. Nội dung nhận xét phải tách thành ĐÚNG 3 CÂU riêng biệt. Giữa mỗi câu có ký tự xuống dòng kép (\\n\\n).`;

            jsonSchema = {
                type: "OBJECT",
                properties: { nhanXet: { type: "STRING" } },
                required: ["nhanXet"]
            };
        } else {
            promptText = `Đóng vai chuyên gia đánh giá tiểu học theo Thông tư 27.
            Mẫu tham chiếu: ${textThamChieu}
            
            BẢNG ĐIỂM TỔNG HỢP CỦA HỌC SINH NÀY:
            [Họ và tên: ${hs.hoTen}] - Kết quả: ${hs.diemSo}
            
            QUY TẮC BẤT DI BẤT DỊCH:
            1. TUYỆT ĐỐI KHÔNG ghi con số điểm hoặc mức T,H,C vào nhận xét.
            2. TUYỆT ĐỐI KHÔNG sử dụng các từ: thầy, cô, giáo viên, học sinh, em, bạn, họ tên, cá nhân, bản thân, người học. Bắt buộc dùng CÂU ẨN CHỦ NGỮ.
            3. Viết NGẮN GỌN, SÚC TÍCH. Phân ra 3 lĩnh vực, mỗi lĩnh vực đúng 1 dòng ngắn.`;

            jsonSchema = {
                type: "OBJECT",
                properties: {
                    nangLucChung: { type: "STRING" },
                    nangLucDacThu: { type: "STRING" },
                    phamChat: { type: "STRING" }
                },
                required: ["nangLucChung", "nangLucDacThu", "phamChat"]
            };
        }

        let retries = 3; 
        let success = false;
        let lastError = "";
        
        while (retries > 0 && !success) {
            try {
                const response = await fetch(apiUrl, {
                    method: "POST", 
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ 
                        contents: [{ parts: [{ text: promptText }] }],
                        generationConfig: { 
                            temperature: 0.3, // Thông số chuẩn nhất cho độ chính xác cao
                            responseMimeType: "application/json",
                            responseSchema: jsonSchema
                        } 
                    })
                });
                
                if (!response.ok) {
                    const errDetail = await response.json();
                    throw new Error(errDetail.error?.message || "Lỗi đường truyền mạng");
                }
                
                const resData = await response.json();
                const aiText = resData.candidates[0].content.parts[0].text;
                
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
                document.getElementById('status').style.color = "#dc3545"; // Đổi màu đỏ khi gặp lỗi
                document.getElementById('status').innerText = `⚠️ Mạng nghẽn tại: ${hs.hoTen}.\nĐang thử lại... Vui lòng đợi...`;
                await delay(4000); 
            }
        }
        
        // HIỂN THỊ LỖI LÊN MÀN HÌNH NGAY LẬP TỨC NẾU THẤT BẠI
        if (!success) {
            console.error("Gặp lỗi mạng với:", hs.hoTen, lastError);
            if (loaiHienTai === "monHoc") {
                hs.nhanXetMonHoc = `[❌ Lỗi xử lý do nghẽn mạng Google API]\n\n[Chi tiết lỗi: ${lastError}]\n\n[Vui lòng xử lý lại em này]`;
            } else { 
                hs.nangLucChung = "[❌ Lỗi xử lý API]"; 
                hs.nangLucDacThu = "[Nghẽn mạng]"; 
                hs.phamChat = "[Vui lòng xử lý lại]"; 
            }
            // Đưa thông báo đỏ nhấp nháy trên Status
            document.getElementById('status').style.color = "#dc3545";
            document.getElementById('status').innerText = `❌ Lỗi phân tích: ${hs.hoTen}.\nỨng dụng tự động chuyển sang học sinh tiếp theo...`;
        }
        
        // Hiện kết quả của học sinh này lên bảng ngay lập tức
        themKetQuaLenBang(hs);
        
        // Thời gian nghỉ an toàn giữa mỗi học sinh
        await delay(8000); 
    }
    
    document.getElementById('status').style.color = "#28a745"; // Màu xanh lá khi thành công
    document.getElementById('status').innerText = "✅ Hoàn tất phân tích dữ liệu một cách Nhanh chóng và Chính xác!";
}

function xuatFileWord() {
    const tieuDeLop = new docx.Paragraph({
        alignment: docx.AlignmentType.CENTER,
        spacing: { after: 300 }, 
        children: [
            new docx.TextRun({
                text: `NHẬN XÉT TỰ ĐỘNG LỚP ${tenLopThucTe}`.toUpperCase(),
                bold: true,
                size: 32, 
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