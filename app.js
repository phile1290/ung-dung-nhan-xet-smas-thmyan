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

async function taoNhanXetVoiAI(apiKey) {
    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;
    
    // TỐI ƯU NHÓM 4 HỌC SINH ĐỂ ĐẢM BẢO CHIỀU SÂU VÀ CHÍNH XÁC TUYỆT ĐỐI
    const BATCH_SIZE = 4;

    for (let i = 0; i < danhSachHocSinh.length; i += BATCH_SIZE) {
        const batch = danhSachHocSinh.slice(i, i + BATCH_SIZE);
        
        document.getElementById('status').innerText = `Ứng dụng đang xử lý chuyên sâu (từ học sinh ${i+1} đến ${i + batch.length} / tổng ${danhSachHocSinh.length}), vui lòng đợi...`;
        
        let thongTinHocSinh = "";
        batch.forEach((hs) => {
            thongTinHocSinh += `[Họ và tên: ${hs.hoTen}]\nKết quả điểm: ${hs.diemSo}\n\n`;
        });

        let promptText = "";

        if (loaiHienTai === "monHoc") {
            promptText = `Đóng vai chuyên gia đánh giá tiểu học theo Thông tư 27. 
            Mẫu tham chiếu: ${textThamChieu}
            
            BẢNG ĐIỂM TỔNG HỢP CỦA DANH SÁCH HỌC SINH NÀY:
            ${thongTinHocSinh}
            
            QUY TẮC BẤT DI BẤT DỊCH (ƯU TIÊN ĐỘ CHÍNH XÁC VÀ CHIỀU SÂU TUYỆT ĐỐI): 
            1. Phải nhận xét đa chiều, sâu sắc, dựa chính xác vào điểm từng môn của từng em. MỖI EM LÀ DUY NHẤT.
            2. TUYỆT ĐỐI KHÔNG ghi các con số điểm (như 7,8,9,10) vào nội dung nhận xét.
            3. TUYỆT ĐỐI KHÔNG sử dụng các từ: thầy, cô, giáo viên, học sinh, em, bạn, họ tên, cá nhân, bản thân, người học. Bắt buộc dùng CÂU ẨN CHỦ NGỮ.
            4. Viết NGẮN GỌN, SÚC TÍCH, đi thẳng vào vấn đề kỹ năng (đọc, tính toán, vẽ, âm nhạc...).
            5. Nội dung nhận xét phải tách thành ĐÚNG 3 CÂU riêng biệt. Giữa mỗi câu có ký tự xuống dòng kép (\\n\\n).
            
            YÊU CẦU ĐẦU RA (BẮT BUỘC TRẢ VỀ CHUẨN JSON ARRAY, KHÔNG KÈM TEXT NGOÀI):
            [
              {
                "hoTen": "Tên học sinh 1",
                "nhanXet": "câu 1.\\n\\ncâu 2.\\n\\ncâu 3."
              },
              {
                "hoTen": "Tên học sinh 2",
                "nhanXet": "câu 1.\\n\\ncâu 2.\\n\\ncâu 3."
              }
            ]`;

        } else {
            promptText = `Đóng vai chuyên gia đánh giá tiểu học theo Thông tư 27.
            Mẫu tham chiếu: ${textThamChieu}
            
            BẢNG ĐIỂM TỔNG HỢP CỦA DANH SÁCH HỌC SINH NÀY:
            ${thongTinHocSinh}
            
            QUY TẮC BẤT DI BẤT DỊCH (ƯU TIÊN ĐỘ CHÍNH XÁC VÀ CHIỀU SÂU TUYỆT ĐỐI):
            1. Đánh giá đa chiều, sâu sắc, chuẩn xác theo dữ liệu điểm của từng em. MỖI EM LÀ DUY NHẤT.
            2. TUYỆT ĐỐI KHÔNG ghi con số điểm hoặc mức T,H,C vào nhận xét.
            3. TUYỆT ĐỐI KHÔNG sử dụng các từ: thầy, cô, giáo viên, học sinh, em, bạn, họ tên, cá nhân, bản thân, người học. Bắt buộc dùng CÂU ẨN CHỦ NGỮ.
            4. Viết NGẮN GỌN, SÚC TÍCH. Phân ra 3 lĩnh vực (Năng lực chung, Năng lực đặc thù, Phẩm chất), mỗi lĩnh vực 1 dòng ngắn.
            
            YÊU CẦU ĐẦU RA (BẮT BUỘC TRẢ VỀ CHUẨN JSON ARRAY, KHÔNG KÈM TEXT NGOÀI):
            [
              {
                "hoTen": "Tên học sinh 1",
                "nangLucChung": "nội dung năng lực chung",
                "nangLucDacThu": "nội dung năng lực đặc thù",
                "phamChat": "nội dung phẩm chất"
              }
            ]`;
        }

        let retries = 4; 
        let success = false;
        let retryWaitTime = 5000; 
        let lastError = "";
        
        while (retries > 0 && !success) {
            try {
                const response = await fetch(apiUrl, {
                    method: "POST", 
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ 
                        contents: [{ parts: [{ text: promptText }] }],
                        generationConfig: { 
                            temperature: 0.4, // Đủ độ mượt và sáng tạo để nhận xét có chiều sâu
                            responseMimeType: "application/json" // Ép chuẩn JSON nhưng không dùng bộ lọc Array Schema cứng nhắc
                        } 
                    })
                });
                
                if (!response.ok) {
                    const errDetail = await response.json();
                    throw new Error(errDetail.error?.message || "Lỗi đường truyền mạng");
                }
                
                const resData = await response.json();
                let aiText = resData.candidates[0].content.parts[0].text;
                
                // Thuật toán làm sạch Markdown an toàn
                aiText = aiText.replace(/```json/gi, "").replace(/```/g, "").trim();
                const ketQuaBatch = JSON.parse(aiText);
                
                ketQuaBatch.forEach(kq => {
                    let hsIndex = danhSachHocSinh.findIndex(h => h.hoTen.toLowerCase().trim() === kq.hoTen.toLowerCase().trim());
                    if (hsIndex !== -1) {
                        if (loaiHienTai === "monHoc") {
                            danhSachHocSinh[hsIndex].nhanXetMonHoc = kq.nhanXet;
                        } else {
                            danhSachHocSinh[hsIndex].nangLucChung = kq.nangLucChung;
                            danhSachHocSinh[hsIndex].nangLucDacThu = kq.nangLucDacThu;
                            danhSachHocSinh[hsIndex].phamChat = kq.phamChat;
                        }
                    }
                });
                
                success = true;
            } catch (err) {
                lastError = err.message;
                retries--;
                console.warn(`Lỗi API: ${lastError}. Đang thử lại sau ${retryWaitTime/1000} giây...`);
                await delay(retryWaitTime); 
                retryWaitTime += 3000; 
            }
        }
        
        // NẾU THẤT BẠI: Hiển thị lỗi rõ ràng chứ tuyệt đối không in câu nhận xét giả mạo
        if (!success) {
            console.error("Gặp lỗi mạng với nhóm này:", lastError);
            batch.forEach(hs => {
                let hsIndex = danhSachHocSinh.findIndex(h => h.hoTen === hs.hoTen);
                if (hsIndex !== -1) {
                    if (loaiHienTai === "monHoc") {
                        danhSachHocSinh[hsIndex].nhanXetMonHoc = `[Lỗi xử lý do nghẽn mạng Google API]\n\n[Hãy kiểm tra lại kết nối]\n\n[Vui lòng bấm Xử lý lại ứng dụng]`;
                    } else { 
                        danhSachHocSinh[hsIndex].nangLucChung = "[Lỗi xử lý API]"; 
                        danhSachHocSinh[hsIndex].nangLucDacThu = "[Nghẽn mạng]"; 
                        danhSachHocSinh[hsIndex].phamChat = "[Vui lòng chạy lại]"; 
                    }
                }
            });
        }
        
        // Thời gian nghỉ 7 giây là lý tưởng nhất khi chạy nhóm 4 học sinh
        await delay(7000); 
    }
    
    document.getElementById('status').innerText = "✅ Hoàn tất phân tích dữ liệu chuyên sâu!";
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