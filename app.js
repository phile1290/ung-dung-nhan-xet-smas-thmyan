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

// BỘ LỌC XỬ LÝ LỖI \n VÀ ÉP CÁCH HÀNG
function xuLyVanBanNhanXet(vanBan) {
    if (!vanBan) return "";
    // Thay thế toàn bộ ký tự \n thô thành dấu xuống dòng thực tế
    let textDaSua = vanBan.replace(/\\n/g, '\n'); 
    // Tách thành mảng, loại bỏ các dòng trống thừa
    let mangCau = textDaSua.split('\n').map(line => line.trim()).filter(line => line !== '');
    // Nối lại bằng đúng 2 dấu xuống dòng để tạo 1 khoảng trắng ở giữa
    return mangCau.join('\n\n'); 
}

async function taoNhanXetVoiAI(apiKey) {
    // SỬ DỤNG MODEL 1.5 FLASH: Miễn phí 1500 lượt/ngày (Giải quyết triệt để lỗi Quota 20 lượt)
    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${apiKey}`;
    
    khoiTaoBang();

    for (let i = 0; i < danhSachHocSinh.length; i++) {
        let hs = danhSachHocSinh[i];
        
        document.getElementById('status').style.color = "#0066cc";
        document.getElementById('status').innerText = `Ứng dụng đang xử lý chuyên sâu: ${hs.hoTen} (${i+1}/${danhSachHocSinh.length})\nVui lòng đợi...`;
        
        let promptText = "";
        let jsonSchema = {};

        if (loaiHienTai === "monHoc") {
            promptText = `Đóng vai chuyên gia đánh giá tiểu học theo Thông tư 27. 
            Mẫu tham chiếu: ${textThamChieu}
            
            BẢNG ĐIỂM TỔNG HỢP CỦA HỌC SINH NÀY:
            [Họ và tên: ${hs.hoTen}] - Kết quả: ${hs.diemSo}
            
            QUY TẮC BẤT DI BẤT DỊCH TỪ GIÁO VIÊN: 
            1. TUYỆT ĐỐI KHÔNG ghi các con số điểm (như 7,8,9,10) vào nội dung nhận xét.
            2. TUYỆT ĐỐI KHÔNG sử dụng đại từ: thầy, cô, giáo viên, học sinh, em, bạn, họ tên, cá nhân, bản thân, người học. Bắt buộc dùng CÂU ẨN CHỦ NGỮ.
            3. Mỗi câu nhận xét phải viết CỰC KỲ CHI TIẾT, có chiều sâu, đánh giá đúng năng lực. MỖI CÂU PHẢI ĐẠT ĐỘ DÀI TỪ 30 ĐẾN 40 TỪ.
            4. Nội dung nhận xét phải được tách thành ĐÚNG 3 CÂU riêng biệt. Giữa mỗi câu dùng ký tự xuống dòng.`;

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
            2. TUYỆT ĐỐI KHÔNG sử dụng đại từ: thầy, cô, giáo viên, học sinh, em, bạn, họ tên, cá nhân, bản thân. Bắt buộc dùng CÂU ẨN CHỦ NGỮ.
            3. Viết phân tích CHI TIẾT, SÂU SẮC, ĐỘ DÀI TỪ 30 ĐẾN 40 TỪ cho mỗi lĩnh vực.
            4. Phân ra 3 lĩnh vực (Năng lực chung, Năng lực đặc thù, Phẩm chất).`;

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
                            temperature: 0.7, 
                            responseMimeType: "application/json",
                            responseSchema: jsonSchema
                        } 
                    })
                });
                
                if (!response.ok) {
                    const errDetail = await response.json();
                    throw new Error(errDetail.error?.message || "Lỗi đường truyền hoặc Quá tải API");
                }
                
                const resData = await response.json();
                const aiText = resData.candidates[0].content.parts[0].text;
                const ketQua = JSON.parse(aiText);
                
                if (loaiHienTai === "monHoc") {
                    hs.nhanXetMonHoc = xuLyVanBanNhanXet(ketQua.nhanXet);
                } else {
                    hs.nangLucChung = ketQua.nangLucChung;
                    hs.nangLucDacThu = ketQua.nangLucDacThu;
                    hs.phamChat = ketQua.phamChat;
                }
                
                success = true;
            } catch (err) {
                lastError = err.message;
                retries--;
                document.getElementById('status').style.color = "#dc3545"; 
                document.getElementById('status').innerText = `⚠️ Mạng nghẽn tại: ${hs.hoTen}.\nĐang thử lại... Vui lòng đợi...`;
                await delay(4000); 
            }
        }
        
        if (!success) {
            console.error("Gặp lỗi mạng với:", hs.hoTen, lastError);
            if (loaiHienTai === "monHoc") {
                hs.nhanXetMonHoc = `[❌ Lỗi xử lý API]\n\n[Chi tiết: ${lastError}]\n\n[Vui lòng xử lý lại em này]`;
            } else { 
                hs.nangLucChung = "[❌ Lỗi API]"; 
                hs.nangLucDacThu = "[Nghẽn mạng]"; 
                hs.phamChat = "[Vui lòng xử lý lại]"; 
            }
            document.getElementById('status').style.color = "#dc3545";
            document.getElementById('status').innerText = `❌ Lỗi phân tích: ${hs.hoTen}.\nỨng dụng tự động chuyển sang học sinh tiếp theo...`;
        }
        
        themKetQuaLenBang(hs);
        await delay(5000); // Rút ngắn thời gian nghỉ an toàn do model 1.5 flash cho phép tần suất gọi cao hơn
    }
    
    document.getElementById('status').style.color = "#28a745"; 
    document.getElementById('status').innerText = "✅ Hoàn tất phân tích dữ liệu chuyên sâu!";
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
            // Tách theo \n\n để chèn đoạn trắng thực tế vào Word
            let parts = hs.nhanXetMonHoc ? hs.nhanXetMonHoc.split('\n\n') : [""];
            let paragraphs = [];
            parts.forEach((p, index) => {
                paragraphs.push(new docx.Paragraph(p));
                // Nếu chưa phải đoạn cuối cùng, chèn thêm 1 đoạn rỗng để tạo khoảng cách hàng chuẩn xác
                if (index < parts.length - 1) {
                    paragraphs.push(new docx.Paragraph("")); 
                }
            });

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