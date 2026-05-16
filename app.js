// CỐ ĐỊNH ĐƯỜNG LINK WEB APP TẠI ĐÂY ĐỂ GIÁO VIÊN KHÔNG PHẢI NHẬP TRÊN GIAO DIỆN
const GAS_WEB_APP_URL = "https://script.google.com/macros/s/AKfycbz-bzcsJlK5NmFJ0GhFDM-rqwHvHxeAzYUwZUHngcWt8Nvbez_OOioA5GKSqUgbAp19Dw/exec";

let danhSachHocSinh = [];
let thongTinThamKhao = {};
let loaiHienTai = "";
let tenTepXuat = "";
let khoiHienTai = "";

// Hàm đọc từng file Excel thành JSON
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
    const files = document.getElementById('fileUpload').files;
    const selectLoai = document.getElementById('loaiBangDiem');
    const selectKhoi = document.getElementById('khoiLop');
    
    loaiHienTai = selectLoai.value;
    khoiHienTai = selectKhoi.value;
    
    const moTaLoai = selectLoai.options[selectLoai.selectedIndex].text;
    const moTaKhoi = selectKhoi.options[selectKhoi.selectedIndex].text;
    tenTepXuat = `NhanXet_${moTaKhoi}_${moTaLoai.replace(/ /g, "_")}.docx`;
    document.getElementById('tenTepHienThi').innerText = tenTepXuat;

    if (!GAS_WEB_APP_URL || GAS_WEB_APP_URL.includes("DÁN_ĐƯỜNG_LINK") || !apiKey || files.length === 0) {
        return alert("Vui lòng cấu hình chính xác đường link Web App trong file mã nguồn, nhập API Key và chọn file bảng điểm!");
    }

    document.getElementById('status').innerText = "Đang kết nối tải dữ liệu cấu hình từ Google Drive...";
    
    try {
        const response = await fetch(GAS_WEB_APP_URL);
        const json = await response.json();
        if (json.status !== "success") throw new Error("Lỗi GAS");
        thongTinThamKhao = json.data;
    } catch (e) {
        return alert("Lỗi kết nối cơ sở dữ liệu qua Web App. Vui lòng kiểm tra lại đường link cố định trong file app.js.");
    }

    document.getElementById('status').innerText = "Đang gộp và phân tích các file bảng điểm...";
    
    danhSachHocSinh = [];
    for (let i = 0; i < files.length; i++) {
        let rawData = await docFileExcel(files[i]);
        for(let j = 1; j < rawData.length; j++) { 
            if(rawData[j][0]) {
                danhSachHocSinh.push({
                    hoTen: rawData[j][0],
                    diemSo: rawData[j].slice(1).join(" | ")
                });
            }
        }
    }
    
    await taoNhanXetVoiAI(apiKey);
}

async function taoNhanXetVoiAI(apiKey) {
    const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=${apiKey}`;
    let mauThamChieu = loaiHienTai === "monHoc" ? thongTinThamKhao.mauMonHoc[khoiHienTai] : thongTinThamKhao.mauNangLuc[khoiHienTai];
    let tt27 = thongTinThamKhao.thongTu27;

    for (let i = 0; i < danhSachHocSinh.length; i++) {
        let hs = danhSachHocSinh[i];
        document.getElementById('status').innerText = `Hệ thống AI đang phân tích: ${hs.hoTen} (${i+1}/${danhSachHocSinh.length})`;
        
        let promptText = "";
        if (loaiHienTai === "monHoc") {
            promptText = `Đóng vai chuyên gia đánh giá tiểu học theo TT27. 
            Tham chiếu Thông tư 27: ${tt27.substring(0,600)}. 
            Mẫu tham khảo (chỉ dùng để tham chiếu văn phong, tuyệt đối KHÔNG copy y nguyên): ${mauThamChieu.substring(0,600)}.
            
            Điểm số và mức đạt được thực tế của học sinh này: "${hs.diemSo}". 
            
            Quy tắc BẤT DI BẤT DỊCH: 
            1. BÁM SÁT VÀO ĐIỂM SỐ của học sinh để đưa ra nhận xét chính xác. Tự do sáng tạo câu từ phù hợp.
            2. Tuyệt đối KHÔNG sử dụng các từ: thầy, cô, giáo viên, học sinh, em, bạn, họ tên.
            3. Viết nhận xét môn học thành ĐÚNG 3 câu. Giữa các câu phải có 1 dòng trống (sử dụng ký tự xuống dòng đôi \\n\\n).
            4. Trả về đúng định dạng JSON: {"nhanXet": "Câu 1.\\n\\nCâu 2.\\n\\nCâu 3."}. Không in bất cứ văn bản nào khác.`;
        } else {
            promptText = `Đóng vai chuyên gia đánh giá tiểu học theo TT27.
            Tham chiếu Thông tư 27: ${tt27.substring(0,600)}. 
            Mẫu tham khảo (chỉ dùng tham chiếu, KHÔNG copy y nguyên): ${mauThamChieu.substring(0,600)}.
            
            Kết quả đánh giá thực tế của học sinh: "${hs.diemSo}".
            
            Quy tắc BẤT DI BẤT DỊCH:
            1. BÁM SÁT VÀO KẾT QUẢ để nhận xét. Tự sáng tạo câu từ.
            2. Tuyệt đối KHÔNG sử dụng: thầy, cô, giáo viên, học sinh, em, bạn, họ tên.
            3. Viết 3 lĩnh vực (Năng lực chung, Năng lực đặc thù, Phẩm chất), mỗi lĩnh vực ĐÚNG 1 dòng ngắn.
            4. Trả về JSON: {"nangLucChung": "...", "nangLucDacThu": "...", "phamChat": "..."}. Không in văn bản nào khác.`;
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
            if (loaiHienTai === "monHoc") hs.nhanXetMonHoc = "Hoàn thành tốt nhiệm vụ.\n\nNắm vững kiến thức.\n\nTích cực học tập.";
            else { hs.nangLucChung = "Giao tiếp tốt"; hs.nangLucDacThu = "Tư duy logic"; hs.phamChat = "Chăm ngoan"; }
        }
    }
    document.getElementById('status').innerText = "Hoàn tất xử lý dữ liệu!";
    hienThiBang();
}

function hienThiBang() {
    document.getElementById('previewArea').style.display = 'block';
    let html = '<table>';
    if (loaiHienTai === 'monHoc') {
        html += '<tr><th style="width: 30%;">Họ và tên</th><th>Đánh giá môn học và HĐGD</th></tr>';
        danhSachHocSinh.forEach(hs => {
            html += `<tr><td>${hs.hoTen}</td><td style="white-space: pre-line;">${hs.nhanXetMonHoc}</td></tr>`;
        });
    } else {
        html += '<tr><th>Họ và tên</th><th>Năng lực chung</th><th>Năng lực đặc thù</th><th>Phẩm chất</th></tr>';
        danhSachHocSinh.forEach(hs => {
            html += `<tr><td>${hs.hoTen}</td><td>${hs.nangLucChung}</td><td>${hs.nangLucDacThu}</td><td>${hs.phamChat}</td></tr>`;
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
            new docx.TableCell({ children: [new docx.Paragraph({text: "Đánh giá môn học và HĐGD", bold: true})] })
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
            new docx.TableCell({ children: [new docx.Paragraph({text: "Năng lực chung", bold: true})] }),
            new docx.TableCell({ children: [new docx.Paragraph({text: "Năng lực đặc thù", bold: true})] }),
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