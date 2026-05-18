let danhSachHocSinh = [];
let loaiHienTai = "";
let tenTepXuat = "";
let tenLopThucTe = "";

// CƠ SỞ DỮ LIỆU TỪ VỰNG NHẬN XÉT (Súc tích, 30-40 từ/câu, không đại từ)
const dataNhanXet = {
    Gioi: {
        cau1: [
            "Nắm vững toàn bộ kiến thức trọng tâm của các môn học cơ bản, thể hiện khả năng tư duy logic và kỹ năng đọc hiểu văn bản xuất sắc. Quá trình tiếp thu bài học diễn ra nhanh chóng, biết cách chắt lọc thông tin và ghi nhớ các khái niệm cốt lõi một cách hệ thống, vững vàng.",
            "Thể hiện sự xuất sắc trong việc nắm bắt kiến thức nền tảng ở các môn học, đặc biệt nổi trội về khả năng tư duy phân tích và giải quyết vấn đề. Năng lực đọc hiểu và tính toán được thực hiện một cách thành thạo, cho thấy sự chủ động và tư duy nhạy bén trong học tập.",
            "Năng lực tiếp thu kiến thức các môn học đạt mức độ rất tốt, duy trì được sự ổn định cao trong suốt quá trình học tập. Kỹ năng tính toán, đọc viết và phân tích dữ liệu được thực hiện một cách trơn tru, thể hiện sự hiểu biết sâu sắc đối với các nội dung trọng điểm."
        ],
        cau2: [
            "Vận dụng linh hoạt các lý thuyết đã học vào việc giải quyết bài tập thực hành, đặc biệt trong các môn Khoa học và Lịch sử - Địa lý. Thao tác thực hành chính xác, nhanh nhẹn, biết cách trình bày sản phẩm học tập một cách khoa học, rõ ràng và có tính thẩm mỹ cao.",
            "Kỹ năng thực hành và áp dụng kiến thức vào giải quyết các tình huống thực tế được thực hiện một cách xuất sắc và mang tính sáng tạo cao. Các bài tập vận dụng được xử lý một cách triệt để, thể hiện sự kết nối nhuần nhuyễn giữa lý thuyết trên lớp và các kỹ năng sống thực tiễn.",
            "Khả năng thực hành các môn học mang tính ứng dụng đạt kết quả rất đáng khích lệ, thao tác kỹ thuật nhanh nhẹn và chuẩn xác. Sản phẩm học tập luôn được hoàn thiện với độ chỉn chu cao, cho thấy sự quan sát tỉ mỉ và năng lực chuyển hóa kiến thức vào thực tiễn một cách hiệu quả."
        ],
        cau3: [
            "Tích cực và chủ động tham gia vào các hoạt động giáo dục trên lớp, thể hiện tinh thần trách nhiệm cao đối với mọi nhiệm vụ được giao. Luôn duy trì thái độ học tập nghiêm túc, sôi nổi phát biểu xây dựng bài và sẵn sàng hỗ trợ bạn bè xung quanh trong quá trình thảo luận nhóm.",
            "Luôn thể hiện thái độ học tập tự giác, tích cực và tràn đầy năng lượng trong các hoạt động phong trào cũng như thực hành trên lớp. Khả năng làm việc nhóm và giao tiếp được phát huy tối đa, góp phần tạo ra môi trường học tập sôi nổi, hiệu quả và lan tỏa tinh thần tích cực.",
            "Tham gia đầy đủ và nhiệt tình vào các hoạt động giáo dục, luôn giữ vững tinh thần ham học hỏi và khám phá những kiến thức mới mẻ. Khả năng tự quản lý công việc và hoàn thành xuất sắc các nhiệm vụ được giao cho thấy sự trưởng thành và ý thức tự giác rất cao trong học tập."
        ]
    },
    Kha: {
        cau1: [
            "Nắm bắt được các kiến thức cơ bản của môn học, đáp ứng đầy đủ các yêu cầu cốt lõi về kỹ năng đọc hiểu và tính toán. Cần tiếp tục duy trì nhịp độ học tập hiện tại và rèn luyện thêm sự cẩn thận để nâng cao hơn nữa độ chính xác trong các bài kiểm tra.",
            "Tiếp thu kiến thức các môn học ở mức độ ổn định, hoàn thành tốt các mục tiêu học tập cơ bản đã được đề ra trong chương trình. Tuy nhiên, cần chú ý phân bổ thời gian hợp lý hơn để rèn luyện chuyên sâu các dạng bài tập phân tích nhằm củng cố vững chắc nền tảng kiến thức.",
            "Kiến thức trọng tâm của các môn học được ghi nhớ và hiểu một cách khá tốt, kỹ năng thực hành các phép tính cơ bản đạt yêu cầu. Để đạt được kết quả cao hơn, cần tích cực rèn luyện thêm khả năng lập luận và mở rộng vốn từ vựng trong các phân môn xã hội."
        ],
        cau2: [
            "Biết cách vận dụng kiến thức lý thuyết vào các bài tập thực hành ở mức độ cơ bản, hoàn thành đầy đủ các sản phẩm học tập được giao. Cần chú ý quan sát tỉ mỉ hơn và rèn luyện thêm thao tác kỹ thuật để các sản phẩm thực hành đạt được độ hoàn thiện và tính thẩm mỹ cao hơn.",
            "Kỹ năng thực hành và giải quyết các vấn đề thực tiễn đạt mức khá, thao tác tương đối nhịp nhàng trong các môn học đòi hỏi kỹ năng vận động. Cần tự tin hơn nữa trong việc đưa ra các ý tưởng sáng tạo và áp dụng linh hoạt các phương pháp khác nhau vào quá trình làm bài thực hành.",
            "Hoàn thành các nhiệm vụ vận dụng thực tiễn ở mức độ tương đối tốt, biết cách liên kết một số kiến thức đã học với môi trường xung quanh. Nên tăng cường tham gia các bài tập nhóm để học hỏi thêm kỹ năng giải quyết tình huống nhanh bén và trau chuốt hơn cho các sản phẩm cuối cùng."
        ],
        cau3: [
            "Có ý thức tham gia vào các hoạt động giáo dục và phong trào của lớp, duy trì được thái độ học tập chuyên cần và nề nếp. Cần mạnh dạn, tự tin hơn nữa trong việc giơ tay phát biểu đóng góp ý kiến để phát huy tối đa năng lực bản thân và tạo ấn tượng nổi bật hơn.",
            "Hoàn thành các nhiệm vụ được phân công trong các hoạt động tập thể một cách đầy đủ và có tinh thần trách nhiệm. Khuyến khích sự chủ động tương tác nhiều hơn với các bạn trong lớp để cải thiện kỹ năng giao tiếp và tăng cường hiệu quả làm việc nhóm trong tương lai.",
            "Thái độ học tập nghiêm túc, luôn cố gắng hoàn thành khối lượng bài vở và các hoạt động giáo dục theo đúng tiến độ yêu cầu. Cần phát huy hơn nữa tinh thần chủ động, tích cực đóng góp ý kiến cá nhân để làm phong phú thêm nội dung các buổi học và rèn luyện sự tự tin."
        ]
    }
};

const dataNangLuc = {
    Gioi: {
        nlc: ["Sở hữu khả năng tự học và giải quyết vấn đề vô cùng xuất sắc, luôn chủ động tìm tòi và khám phá các phương pháp học tập mới mẻ, mang lại hiệu quả cao.", "Giao tiếp tự tin, mạch lạc và có khả năng phối hợp làm việc nhóm cực kỳ hiệu quả, luôn giữ vai trò nòng cốt trong các hoạt động tập thể."],
        nldt: ["Tư duy logic nhạy bén, thể hiện năng lực ngôn ngữ và tính toán vượt trội, giải quyết các tình huống học thuật phức tạp một cách nhanh chóng và chính xác.", "Phát huy tối đa năng lực thẩm mỹ và thể chất, vận dụng linh hoạt kiến thức khoa học vào thực tiễn, tạo ra các sản phẩm sáng tạo mang tính ứng dụng cao."],
        pc: ["Thể hiện rõ tinh thần trách nhiệm, luôn trung thực, đoàn kết và sẵn sàng chia sẻ, giúp đỡ mọi người xung quanh, là tấm gương sáng về đạo đức.", "Chăm chỉ, kỷ luật và có ý thức cộng đồng rất cao, luôn tôn trọng nội quy và thể hiện tình yêu thương, chan hòa với bạn bè trong mọi hoạt động."]
    },
    Kha: {
        nlc: ["Có ý thức tự giác trong học tập, hoàn thành tốt các nhiệm vụ được giao, biết cách phân bổ thời gian hợp lý để giải quyết các vấn đề cơ bản.", "Giao tiếp hòa nhã, biết lắng nghe và tôn trọng ý kiến của người khác, phối hợp nhịp nhàng với các thành viên trong nhóm khi thực hiện nhiệm vụ."],
        nldt: ["Nắm vững kiến thức nền tảng, có khả năng tính toán và diễn đạt ngôn ngữ tương đối trôi chảy, đáp ứng tốt các yêu cầu cơ bản của môn học.", "Biết cách áp dụng kiến thức vào các bài tập thực hành một cách ổn định, tuy nhiên cần rèn luyện thêm sự cẩn thận để nâng cao độ chính xác."],
        pc: ["Có thái độ hòa đồng, thân thiện với bạn bè, luôn chấp hành đúng các quy định chung và thể hiện sự nỗ lực, cố gắng vươn lên trong học tập.", "Biết giữ gìn vệ sinh chung, có ý thức bảo vệ tài sản công cộng và thể hiện lòng nhân ái, sẵn sàng hỗ trợ khi các bạn gặp khó khăn trong bài vở."]
    }
}

// Hàm lấy phần tử ngẫu nhiên trong mảng
function getRan(arr) { return arr[Math.floor(Math.random() * arr.length)]; }

// Hàm đánh giá mức độ điểm (Giỏi / Khá) để chọn kho từ vựng
function danhGiaDiem(diemStr) {
    let diem = diemStr.toUpperCase();
    if (diem.includes('10') || diem.includes('9') || diem.includes('T')) return 'Gioi';
    return 'Kha'; // Đơn giản hóa thành 2 mức để luôn có nhận xét tích cực, mang tính khích lệ
}

async function batDauXuLy() {
    const filesExcel = document.getElementById('fileUpload').files;
    const selectLoai = document.getElementById('loaiBangDiem');
    
    loaiHienTai = selectLoai.value;
    const tenLopInput = document.getElementById('tenLop');
    tenLopThucTe = tenLopInput ? tenLopInput.value.trim() : "";
    
    const moTaLoai = selectLoai.options[selectLoai.selectedIndex].text;
    tenTepXuat = `NhanXet_${tenLopThucTe ? tenLopThucTe + "_" : ""}${moTaLoai.replace(/ /g, "_")}.docx`;
    document.getElementById('tenTepHienThi').innerText = tenTepXuat;

    if (filesExcel.length === 0 || !tenLopThucTe) {
        return alert("Vui lòng nhập Tên lớp và tải lên file Bảng điểm Excel!");
    }

    document.getElementById('status').innerText = "Đang phân tích dữ liệu ngoại tuyến tốc độ cao...";
    
    let hocSinhTongHop = {};
    
    // Đọc Excel
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
                            !tenLower.includes("giáo viên") && !tenLower.includes("họ và tên") && !tenLower.includes("người lập")) {
                            let diemSoCuaMon = [];
                            for (let c = cotHoTen + 1; c < rawData[r].length; c++) {
                                if (rawData[r][c] !== undefined && rawData[r][c] !== "") diemSoCuaMon.push(rawData[r][c]);
                            }
                            if (diemSoCuaMon.length > 0) {
                                if (!hocSinhTongHop[tenHS]) hocSinhTongHop[tenHS] = [];
                                hocSinhTongHop[tenHS].push(diemSoCuaMon.join(""));
                            }
                        }
                    }
                }
            }
        });
    }
    
    danhSachHocSinh = Object.keys(hocSinhTongHop).map(ten => ({
        hoTen: ten,
        diemSo: hocSinhTongHop[ten].join(" ")
    }));
    
    if(danhSachHocSinh.length === 0) {
        return alert("Không tìm thấy dữ liệu học sinh hợp lệ!");
    }

    // TẠO NHẬN XÉT BẰNG THUẬT TOÁN (Tốc độ 1 giây)
    danhSachHocSinh.forEach(hs => {
        let mucDo = danhGiaDiem(hs.diemSo);
        
        if (loaiHienTai === "monHoc") {
            // Ghép 3 câu, cách nhau bởi 2 dấu \n để tạo 1 dòng trống
            hs.nhanXetMonHoc = getRan(dataNhanXet[mucDo].cau1) + "\n\n" + getRan(dataNhanXet[mucDo].cau2) + "\n\n" + getRan(dataNhanXet[mucDo].cau3);
        } else {
            hs.nangLucChung = getRan(dataNangLuc[mucDo].nlc);
            hs.nangLucDacThu = getRan(dataNangLuc[mucDo].nldt);
            hs.phamChat = getRan(dataNangLuc[mucDo].pc);
        }
    });

    document.getElementById('status').style.color = "#28a745"; 
    document.getElementById('status').innerText = "✅ Đã xử lý xong toàn bộ học sinh thành công!";
    hienThiBang();
}

function hienThiBang() {
    document.getElementById('previewArea').style.display = 'block';
    const container = document.getElementById('tableContainer');
    
    let html = '<table id="bangKetQua">';
    if (loaiHienTai === 'monHoc') {
        html += '<thead><tr><th style="width: 25%;">Họ và tên</th><th>Nhận xét đánh giá môn học và HĐGD</th></tr></thead><tbody>';
        danhSachHocSinh.forEach(hs => {
            html += `<tr><td style="font-weight: bold; color: #004085;">${hs.hoTen}</td><td style="white-space: pre-line;">${hs.nhanXetMonHoc}</td></tr>`;
        });
    } else {
        html += '<thead><tr><th style="width: 20%;">Họ và tên</th><th>Đánh giá năng lực chung</th><th>Đánh giá năng lực đặc thù</th><th>Phẩm chất</th></tr></thead><tbody>';
        danhSachHocSinh.forEach(hs => {
            html += `<tr><td style="font-weight: bold; color: #004085;">${hs.hoTen}</td><td>${hs.nangLucChung}</td><td>${hs.nangLucDacThu}</td><td>${hs.phamChat}</td></tr>`;
        });
    }
    html += '</tbody></table>';
    container.innerHTML = html;
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
            // Tách bằng \n\n và thêm Paragraph trống để tạo khoảng cách trong Word
            let parts = hs.nhanXetMonHoc.split('\n\n');
            let paragraphs = [];
            parts.forEach((p, index) => {
                paragraphs.push(new docx.Paragraph(p));
                if (index < parts.length - 1) paragraphs.push(new docx.Paragraph("")); 
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