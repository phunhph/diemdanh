const express = require('express');
const cors = require('cors');
const xlsx = require('xlsx');
const bodyParser = require('body-parser');
const fs = require('fs');

const app = express();
const port = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(bodyParser.json());

// Đọc dữ liệu từ file Excel
const readExcelData = () => {
    try {
        const workbook = xlsx.readFile('./data.xlsx');
        const sheetNameList = workbook.SheetNames;
        return xlsx.utils.sheet_to_json(workbook.Sheets[sheetNameList[0]]);
    } catch (error) {
        console.error('Error reading Excel:', error);
        throw new Error('Không thể đọc dữ liệu từ file Excel.');
    }
};

// Ghi dữ liệu vào file Excel (cập nhật)
const writeExcelData = (data) => {
    try {
        const newWorkbook = xlsx.utils.book_new();
        const newSheet = xlsx.utils.json_to_sheet(data);
        xlsx.utils.book_append_sheet(newWorkbook, newSheet, 'Sheet1');
        xlsx.writeFile(newWorkbook, './data.xlsx');
    } catch (error) {
        console.error('Error writing Excel:', error);
        throw new Error('Không thể ghi dữ liệu vào file Excel.');
    }
};

// Endpoint để lấy dữ liệu thí sinh
app.get('/api/students', (req, res) => {
    try {
        const data = readExcelData();
        res.json(data);
    } catch (error) {
        res.status(500).json({ message: error.message });
    }
});

// Endpoint để cập nhật dữ liệu thí sinh
app.post('/api/students', (req, res) => {
    const newStudent = req.body;
    const existingData = readExcelData();
    
    // Thêm hoặc cập nhật sinh viên
    const existingStudentIndex = existingData.findIndex(student => student.mssv === newStudent.mssv);
    
    if (existingStudentIndex >= 0) {
        // Cập nhật sinh viên hiện có
        existingData[existingStudentIndex] = { ...existingData[existingStudentIndex], ...newStudent };
    } else {
        // Thêm sinh viên mới vào mảng dữ liệu
        existingData.push(newStudent);
    }

    // Ghi lại vào file Excel
    try {
        writeExcelData(existingData);
        res.status(201).json(newStudent);
    } catch (error) {
        res.status(500).json({ message: 'Có lỗi xảy ra khi ghi dữ liệu vào file.' });
    }
});

// Endpoint để cập nhật trạng thái điểm danh của thí sinh
app.post('/api/students/attendance', (req, res) => {
    const attendanceData = req.body; // Dữ liệu điểm danh từ client

    // Kiểm tra nếu attendanceData có dữ liệu
    if (!attendanceData || !Array.isArray(attendanceData)) {
        return res.status(400).json({ message: 'Dữ liệu không hợp lệ' });
    }

    let existingData;
    try {
        existingData = readExcelData();
    } catch (error) {
        return res.status(500).json({ message: 'Có lỗi khi đọc dữ liệu.' });
    }

    // Cập nhật trạng thái điểm danh
    attendanceData.forEach(attendance => {
        if (!attendance.mssv || attendance.status === undefined) {
            return res.status(400).json({ message: 'Thông tin sinh viên không đầy đủ' });
        }

        const existingStudentIndex = existingData.findIndex(student => student.mssv === attendance.mssv);
        if (existingStudentIndex >= 0) {
            // Cập nhật trạng thái điểm danh
            existingData[existingStudentIndex].status = attendance.status;
            existingData[existingStudentIndex].note = attendance.note;
        } else {
            console.warn(`Sinh viên với MSSV ${attendance.mssv} không tồn tại.`);
        }
    });

    // Ghi lại vào file Excel
    try {
        writeExcelData(existingData);
        res.status(200).json({ message: 'Trạng thái điểm danh đã được cập nhật!' });
    } catch (error) {
        console.error('Error writing to Excel:', error);
        return res.status(500).json({ message: 'Có lỗi xảy ra khi ghi dữ liệu vào file.' });
    }
});

// Endpoint để xuất file Excel với trạng thái điểm danh
app.post('/api/students/export', (req, res) => {
    const attendanceData = req.body; // Dữ liệu điểm danh từ client
    const outputData = attendanceData.map(student => ({
        mssv: student.mssv,
        status: student.status ? 'Đã điểm danh' : 'Chưa điểm danh',
    }));

    // Tạo file Excel
    const newWorkbook = xlsx.utils.book_new();
    const newSheet = xlsx.utils.json_to_sheet(outputData);
    xlsx.utils.book_append_sheet(newWorkbook, newSheet, 'Điểm danh');

    // Đặt tên file và ghi vào file
    const filePath = `attendance-${Date.now()}.xlsx`;
    xlsx.writeFile(newWorkbook, filePath);

    // Gửi file về client
    res.download(filePath, err => {
        if (err) {
            res.status(500).send('Có lỗi xảy ra khi tải file!');
        }
        // Xóa file sau khi gửi (tuỳ chọn)
        fs.unlink(filePath, (err) => {
            if (err) console.error(err);
        });
    });
});

// Bắt đầu server
app.listen(port, () => {
    console.log(`Server is running on port ${port}`);
});
