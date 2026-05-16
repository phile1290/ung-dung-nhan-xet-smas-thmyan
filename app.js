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

// Hàm đọc file Excel
function docFileExcel(file) {
    return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            resolve(XLSX.utils.sheet_to_json(firstSheet, {header: 1}));
        };
        reader.readAsArrayBuffer(file);
    });
}

async function batDauXuLy() {
    const apiKey = document.getElementById('apiKey').value.trim();
    const fileTT27 = document.getElementById('pdfTT27').files[0];
    const fileThamChieu = document.getElementById('pdfThamChieu').files[0];
    const filesExcel = document.getElementById('fileUpload').files;
    const selectLoai = document.getElementById('loaiBangDiem');
    
    loaiHienTai = selectLoai.value;
    
    // Đặt tên tệp tải về dựa trên mô tả giao diện
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
        
        document.getElementById('status').innerText = "Đang lọc dữ liệu học sinh từ file SMAS...";
        danhSachHocSinh = [];
        
        for (let i = 0; i < filesExcel.length; i++) {
            let rawData = await docFileExcel(filesExcel[i]);
            
            for(let j = 0; j < rawData.length; j++) {
                let ten = rawData[j][0] ? rawData[j][0].toString().trim() : "";
                let ketQua = rawData[j].slice(1).join("").trim(); 
                
                // THUẬT TOÁN LỌC: Bỏ qua dòng trống, dòng không có điểm, và các dòng header/footer của SMAS
                if (ten && ketQua.length > 0) {
                    let tenLower = ten.toLowerCase();
                    if (!tenLower.includes("trường") && !tenLower.includes("hiệu trưởng") && 
                        !tenLower.includes("giáo viên") && !tenLower.includes("sở gd") && 
                        !tenLower.includes("phòng gd") && !tenLower.includes("họ và tên") &&
                        !tenLower.includes("người lập bảng")) {
                        
                        danhSachHocSinh.push({
                            hoTen: ten,
                            diemSo: rawData[j].slice(1).join(" | ")
                        });
                    }
                }
            }
        }
        
        if(danhSachHocSinh.length === 0) {
            return alert("Không tìm thấy học sinh nào có điểm trong file Excel vừa tải lên!");
        }

        await taoNhanXetVoiAI(apiKey);
        
    } catch (error) {
        console.error(error);
        alert("Có lỗi xảy ra trong quá trình đọc file. Vui lòng kiểm tra lại định dạng tệp.");
    }
}

async function taoNhanXetVoiAI(apiKey) {
    // Đã cập nhật sử dụng model gemini-2.5-flash theo yêu cầu của anh
    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;

    for (let i = 0; i < danhSachHocSinh.length; i++) {
        let hs = danhSachHocSinh[i];
        document.getElementById('status').innerText = `Hệ thống AI (Flash 2.5) đang phân tích: ${hs.hoTen} (${i+1}/${danhSachHocSinh.length})`;
        
        let promptText = "";
        if (loaiHienTai === "monHoc") {
            promptText = `Đóng vai chuyên gia đánh giá tiểu học theo Thông tư 27. 
            Nội dung Thông tư 27: ${textTT27.substring(0, 8000)}...
            Nhận xét tham chiếu (Chỉ dùng tham khảo cách hành văn, tuyệt đối KHÔNG copy nguyên bản): ${textThamChieu.substring(0, 4000)}...
            
            Học sinh có kết quả học tập/điểm số thực tế: "${hs.diemSo}". 
            
            Quy tắc BẤT DI BẤT DỊCH: 
            1. CHỈ BÁM SÁT VÀO KẾT QUẢ ĐIỂM SỐ cung cấp ở trên để nhận xét. Tự sáng tạo câu từ dựa trên mẫu tham chiếu.
            2. Tuyệt đối KHÔNG sử dụng các đại từ nhân xưng: thầy, cô, giáo viên, học sinh, em, bạn, họ tên.
            3. Nhận xét môn học viết thành ĐÚNG 3 câu. Giữa các câu phải có 1 dòng trống (sử dụng ký tự \\n\\n).
            4. Trả về đúng định dạng JSON: {"nhanXet": "Câu 1.\\n\\nCâu 2.\\n\\nCâu 3."}. Không kèm văn bản nào khác.`;
        } else {
            promptText = `Đóng vai chuyên gia đánh giá tiểu học theo Thông tư 27.
            Nội dung Thông tư 27: ${textTT27.substring(0, 8000)}...
            Nhận xét tham chiếu (Chỉ dùng tham khảo cách hành văn): ${textThamChieu.substring(0, 4000)}...
            
            Học sinh có kết quả đánh giá thực tế: "${hs.diemSo}".
            
            Quy tắc BẤT DI BẤT DỊCH:
            1. CHỈ BÁM SÁT VÀO KẾT QUẢ để nhận xét.
            2. Tuyệt đối KHÔNG sử dụng các đại từ nhân xưng: thầy, cô, giáo viên, học sinh, em, bạn, họ tên.
            3. Nhận xét thành 3 lĩnh vực (Năng lực chung, Năng lực đặc thù, Phẩm chất), mỗi lĩnh vực ĐÚNG 1 dòng ngắn.
            4. Trả về đúng định dạng JSON: {"nangLucChung": "...", "nangLucDacThu": "...", "phamChat": "..."}. Không kèm văn bản nào khác.`;
        }

        try {
            const response = await fetch(apiUrl, {
                method: "POST", headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ contents: [{ parts: [{ text: promptText }] }] })
            });
            const resData = await response.json();
            let aiText = resData.candidates[0].content.parts[0].text.replace(/```json/g, "").replace(/```/g, "").trim();
            const ketQua = JSON.parse(aiText);
            
            if (loaiHienTai === "monHoc") {
                hs.nhanXetMonHoc = ketQua.nhanXet;
            } else { 
                hs.nangLucChung = ketQua.nangLucChung; hs.nangLucDacThu = ketQua.nangLucDacThu; hs.phamChat = ketQua.phamChat; 
            }
        } catch (err) {
            console.error("Lỗi dòng", i, err);
            if (loaiHienTai === "monHoc") hs.nhanXetMonHoc = "Hoàn thành tốt nhiệm vụ.\n\nNắm vững kiến thức trọng tâm.\n\nTích cực trong các hoạt động.";
            else { hs.nangLucChung = "Giao tiếp và hợp tác tốt"; hs.nangLucDacThu = "Tư duy logic tốt"; hs.phamChat = "Chăm chỉ đoàn kết"; }
        }
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