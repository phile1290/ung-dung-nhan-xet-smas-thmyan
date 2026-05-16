// Cấu hình worker cho thư viện PDF.js
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';

let danhSachHocSinh = [];
let textTT27 = "";
let textThamChieu = "";
let loaiHienTai = "";
let tenTepXuat = "";
let tenLopThucTe = ""; 

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

// Lấy delay để tránh bị lỗi giới hạn API (Rate Limit) gây ra lỗi trùng lặp
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
    
    // Đặt tên tệp xuất ra có chứa tên lớp
    tenTepXuat = `NhanXet_${tenLopThucTe ? tenLopThucTe + "_" : ""}${moTaLoai.replace(/ /g, "_")}.docx`;
    document.getElementById('tenTepHienThi').innerText = tenTepXuat;

    if (!apiKey || !fileTT27 || !fileThamChieu || filesExcel.length === 0 || !tenLopThucTe) {
        return alert("Vui lòng nhập API Key, Tên lớp và tải lên đầy đủ 3 loại tệp dữ liệu!");
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
        document.getElementById('status').innerText = `Hệ thống AI đang phân tích chi tiết: ${hs.hoTen} (${i+1}/${danhSachHocSinh.length})\nĐang áp dụng thời gian chờ 10 giây để đảm bảo an toàn tối đa...`;
        
        let promptText = "";
        if (loaiHienTai === "monHoc") {
            promptText = `Đóng vai chuyên gia đánh giá tiểu học theo Thông tư 27. 
            Nội dung Thông tư 27: ${textTT27.substring(0, 4000)}...
            Nhận xét tham chiếu: ${textThamChieu.substring(0, 3000)}...
            
            Đây là BẢNG ĐIỂM TỔNG HỢP NHIỀU MÔN của học sinh: "${hs.diemSo}". 
            
            Quy tắc BẤT DI BẤT DỊCH: 
            1. BÁM SÁT KẾT QUẢ ĐIỂM SỐ từng môn. TỰ DO SÁNG TẠO nhận xét, đảm bảo TÍNH DUY NHẤT cho học sinh này, KHÔNG TRÙNG LẶP.
            2. Nhận xét CHI TIẾT và RÕ RÀNG (tốc độ đọc, hiểu bài, viết câu, tính toán...).
            3. TUYỆT ĐỐI KHÔNG sử dụng các từ: thầy, cô, giáo viên, học sinh, em, bạn, họ tên, cá nhân, bản thân, người học. Sử dụng câu ẩn chủ ngữ.
            4. Nhận xét thành ĐÚNG 3 CÂU SÚC TÍCH. Giữa mỗi câu PHẢI CÓ một dòng trống (dùng ký tự \\n\\n).
            5. Trả về đúng định dạng JSON: {"nhanXet": "Câu 1.\\n\\nCâu 2.\\n\\nCâu 3."}`;
        } else {
            promptText = `Đóng vai chuyên gia đánh giá tiểu học theo Thông tư 27.
            Nội dung TT27: ${textTT27.substring(0, 4000)}...
            Nhận xét tham chiếu: ${textThamChieu.substring(0, 3000)}...
            
            Kết quả tổng hợp các môn của học sinh: "${hs.diemSo}".
            
            Quy tắc BẤT DI BẤT DỊCH:
            1. BÁM SÁT KẾT QUẢ TỔNG HỢP để nhận xét Năng lực và Phẩm chất.
            2. Mỗi học sinh là DUY NHẤT. Sáng tạo từ vựng, KHÔNG TRÙNG LẶP.
            3. TUYỆT ĐỐI KHÔNG sử dụng các từ: thầy, cô, giáo viên, học sinh, em, bạn, họ tên, cá nhân, bản thân, người học. Sử dụng câu ẩn chủ ngữ.
            4. Viết thành 3 lĩnh vực (Năng lực chung, Năng lực đặc thù, Phẩm chất), mỗi lĩnh vực 1 dòng chi tiết.
            5. Trả về định dạng JSON: {"nangLucChung": "...", "nangLucDacThu": "...", "phamChat": "..."}`;
        }

        let retries = 4; 
        let success = false;
        let waitTime = 10000; // Thời gian chờ mặc định khi lỗi mạng tăng lên 10 giây
        
        while (retries > 0 && !success) {
            try {
                const response = await fetch(apiUrl, {
                    method: "POST", headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ 
                        contents: [{ parts: [{ text: promptText }] }],
                        generationConfig: { 
                            temperature: 0.8,
                            responseMimeType: "application/json" 
                        } 
                    })
                });
                
                if (!response.ok) throw new Error("API Limit");
                
                const resData = await response.json();
                const aiText = resData.candidates[0].content.parts[0].text;
                
                const ketQua = JSON.parse(aiText);
                
                if (loaiHienTai === "monHoc") {
                    hs.nhanXetMonHoc = ketQua.nhanXet;
                } else { 
                    hs.nangLucChung = ketQua.nangLucChung; hs.nangLucDacThu = ketQua.nangLucDacThu; hs.phamChat = ketQua.phamChat; 
                }
                success = true;
            } catch (err) {
                retries--;
                await delay(waitTime); 
                waitTime += 2000; 
                
                if (retries === 0) {
                    console.error("Lỗi dòng học sinh", hs.hoTen);
                    if (loaiHienTai === "monHoc") hs.nhanXetMonHoc = "Tiếp thu bài tốt.\n\nHoàn thành các bài tập thực hành.\n\nCần chú ý cẩn thận hơn trong học tập.";
                    else { hs.nangLucChung = "Tự giác trong học tập"; hs.nangLucDacThu = "Vận dụng kiến thức ổn định"; hs.phamChat = "Đoàn kết hòa đồng"; }
                }
            }
        }
        
        // THỜI GIAN CHỜ VÀNG: 10 giây (10000ms) giữa mỗi học sinh theo yêu cầu
        await delay(10000); 
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