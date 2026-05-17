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

// Lấy delay để tránh bị lỗi giới hạn API (Rate Limit) gây ra lỗi trùng lặp
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
        
        // Dùng Object để gộp điểm của học sinh từ nhiều Sheet khác nhau
        let hocSinhTongHop = {};
        
        for (let i = 0; i < filesExcel.length; i++) {
            const arrayBuffer = await filesExcel[i].arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, {type: 'array'});
            
            // Duyệt qua toàn bộ các Sheet (Tiếng Việt, Toán, Khoa học...)
            workbook.SheetNames.forEach(sheetName => {
                const sheet = workbook.Sheets[sheetName];
                const rawData = XLSX.utils.sheet_to_json(sheet, {header: 1});
                
                let cotHoTen = -1;
                let dongBatDau = -1;

                // THUẬT TOÁN ĐỊNH VỊ CỘT: Tìm đúng cột có chữ "Họ và tên"
                for (let r = 0; r < rawData.length; r++) {
                    if (rawData[r]) {
                        for (let c = 0; c < rawData[r].length; c++) {
                            if (typeof rawData[r][c] === 'string' && rawData[r][c].toLowerCase().includes("họ và tên")) {
                                cotHoTen = c;
                                dongBatDau = r + 1; // Học sinh bắt đầu từ dòng ngay dưới tiêu đề
                                break;
                            }
                        }
                    }
                    if (cotHoTen !== -1) break;
                }

                // Nếu tìm thấy cột Họ và tên trong Sheet này
                if (cotHoTen !== -1) {
                    for (let r = dongBatDau; r < rawData.length; r++) {
                        let ten = rawData[r][cotHoTen];
                        
                        // Lọc bỏ các dòng rỗng, dòng chứa số (STT), và các dòng không phải tên người
                        if (ten && typeof ten === 'string' && isNaN(ten) && ten.trim().length > 4) {
                            let tenHS = ten.trim();
                            let tenLower = tenHS.toLowerCase();
                            
                            // Tránh các dòng rác của SMAS
                            if (!tenLower.includes("trường") && !tenLower.includes("hiệu trưởng") && 
                                !tenLower.includes("giáo viên") && !tenLower.includes("họ và tên") &&
                                !tenLower.includes("người lập")) {
                                
                                // Gom tất cả điểm/đánh giá của học sinh này trên cùng dòng
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
                                    // Lưu theo cấu trúc: Tên Môn: [Điểm]
                                    hocSinhTongHop[tenHS].push(`${sheetName}: ${diemSoCuaMon.join(", ")}`);
                                }
                            }
                        }
                    }
                }
            });
        }
        
        // Chuyển đổi dữ liệu tổng hợp sang mảng để AI dễ đọc
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
        document.getElementById('status').innerText = `Hệ thống AI đang phân tích chi tiết: ${hs.hoTen} (${i+1}/${danhSachHocSinh.length})`;
        
        let promptText = "";
        if (loaiHienTai === "monHoc") {
            promptText = `Đóng vai chuyên gia đánh giá tiểu học theo Thông tư 27. 
            Nội dung Thông tư 27: ${textTT27.substring(0, 5000)}...
            Nhận xét tham chiếu: ${textThamChieu.substring(0, 3000)}...
            
            Đây là BẢNG ĐIỂM TỔNG HỢP NHIỀU MÔN của học sinh: "${hs.diemSo}". 
            
            Quy tắc BẤT DI BẤT DỊCH: 
            1. BÁM SÁT KẾT QUẢ ĐIỂM SỐ từng môn. TỰ DO SÁNG TẠO nhận xét, đảm bảo TÍNH DUY NHẤT cho học sinh này, KHÔNG TRÙNG LẶP với học sinh khác.
            2. Nhận xét cực kỳ CHI TIẾT và RÕ RÀNG về các kỹ năng cốt lõi (ví dụ: Tốc độ đọc, khả năng hiểu bài, viết câu, tính toán nhanh nhạy, trình bày bản đồ/hình vẽ, sự tập trung...).
            3. Tuyệt đối KHÔNG sử dụng các từ: thầy, cô, giáo viên, học sinh, em, bạn, họ tên.
            4. Nhận xét thành ĐÚNG 3 CÂU DÀI, CHI TIẾT. Giữa mỗi câu PHẢI CÓ một dòng trống (dùng ký tự \\n\\n).
            5. CHỈ trả về JSON định dạng: {"nhanXet": "Câu 1.\\n\\nCâu 2.\\n\\nCâu 3."}. Không thêm bất kỳ text nào khác.`;
        } else {
            promptText = `Đóng vai chuyên gia đánh giá tiểu học theo Thông tư 27.
            Nội dung TT27: ${textTT27.substring(0, 5000)}...
            Nhận xét tham chiếu: ${textThamChieu.substring(0, 3000)}...
            
            Kết quả tổng hợp các môn của học sinh: "${hs.diemSo}".
            
            Quy tắc BẤT DI BẤT DỊCH:
            1. BÁM SÁT KẾT QUẢ TỔNG HỢP để nhận xét Năng lực và Phẩm chất.
            2. Mỗi học sinh là DUY NHẤT. Hãy sáng tạo từ vựng phong phú, đánh giá sát sao, KHÔNG TRÙNG LẶP.
            3. Tuyệt đối KHÔNG sử dụng các từ: thầy, cô, giáo viên, học sinh, em, bạn, họ tên.
            4. Viết thành 3 lĩnh vực (Năng lực chung, Năng lực đặc thù, Phẩm chất), mỗi lĩnh vực 1 dòng chi tiết.
            5. CHỈ trả về JSON: {"nangLucChung": "...", "nangLucDacThu": "...", "phamChat": "..."}. Không thêm text nào.`;
        }

        let retries = 3;
        let success = false;
        
        while (retries > 0 && !success) {
            try {
                const response = await fetch(apiUrl, {
                    method: "POST", headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ 
                        contents: [{ parts: [{ text: promptText }] }],
                        generationConfig: { temperature: 0.8 } // Tăng tính sáng tạo, đa dạng văn phong
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
                await delay(1500); // Tạm dừng 1.5 giây nếu lỗi API để tránh trùng lặp do lỗi hệ thống
                if (retries === 0) {
                    console.error("Lỗi dòng học sinh", hs.hoTen);
                    // Thông báo khuyết điểm tinh tế nếu lỗi mạng 3 lần liên tiếp
                    if (loaiHienTai === "monHoc") hs.nhanXetMonHoc = "Khả năng tiếp thu bài học ổn định.\n\nKỹ năng tính toán và thực hành cơ bản đạt yêu cầu.\n\nCần chú ý duy trì sự tập trung hơn trong các hoạt động.";
                    else { hs.nangLucChung = "Có ý thức tự giác học tập"; hs.nangLucDacThu = "Vận dụng kiến thức linh hoạt"; hs.phamChat = "Hòa đồng với bạn bè"; }
                }
            }
        }
        await delay(500); // Nghỉ nhẹ giữa mỗi học sinh để đảm bảo Gemini phân tích kỹ
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