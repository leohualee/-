// 取得所有需要的 DOM 元素 
const classNameInput = document.getElementById('classNameInput');
const fileInput = document.getElementById('fileInput');
const importBtn = document.getElementById('importBtn');
const seatingChartGrid = document.querySelector('.seating-chart-grid');
const scoreSummary = document.getElementById('scoreSummary');
const downloadExcelBtn = document.getElementById('downloadExcelBtn');

// 儲存學生資料的物件，鍵是座位ID，值是學生資料
let students = {};
let currentClassName = '';

// 國字數字對應
const chineseNumerals = ['零', '一', '二', '三', '四', '五', '六'];

// 作品評級的分數對應
const projectGradeScores = {
    'A': 100,
    'B': 90,
    'C': 80,
    'D': 70,
    'E': 60,
    '作品': 0,
    '假': 0
};

// 作品調整分數的對應
const projectModifierScores = {
    '+': 5,
    '-': -5,
    '無': 0
};

// 監聽班級名稱輸入框的輸入事件
classNameInput.addEventListener('input', () => {
    const value = classNameInput.value.trim();
    if (!isNaN(parseInt(value)) && value.length > 0) {
        classNameInput.value = parseInt(value) + '年班';
    } else {
        classNameInput.value = value; // 允許輸入其他文字
    }
    currentClassName = classNameInput.value || '未命名班級';
});

// 事件監聽器 - 匯入按鈕
importBtn.addEventListener('click', () => {
    const files = fileInput.files;
    if (files.length === 0) {
        alert('請選擇一個 Excel 或 CSV 檔案！');
        return;
    }
    const file = files[0];
    const fileName = file.name.toLowerCase();

    // 在這裡加入解析檔名的邏輯
    const classNameFromFile = extractClassName(file.name);
    if (classNameFromFile) {
        classNameInput.value = classNameFromFile;
    }
    
    if (fileName.endsWith('.csv')) {
        Papa.parse(file, {
            complete: function(results) {
                processImportedData(results.data);
            },
            header: false
        });
    } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName], { header: 1 });
                processImportedData(worksheet);
            } catch (error) {
                console.error('匯入檔案時發生錯誤:', error);
                alert('無法讀取檔案，請確認它是有效的 Excel 檔案且格式正確。');
            }
        };
        reader.readAsArrayBuffer(file);
    } else {
        alert('不支援的檔案格式，請匯入 .xlsx、.xls 或 .csv 檔案！');
    }
});

// 新增一個解析檔名的函數
function extractClassName(fileName) {
    const baseName = fileName.split('.')[0];
    // 正規表達式匹配各種可能的班級格式
    // 範例：1年1班, 1-1, 101, 701, 7年1班
    let match = baseName.match(/(\d+年\d+班|\d+-\d+|\d{3}|\d+)/);
    
    if (match && match[0]) {
        return match[0].replace(/-/g, '年'); // 將 1-1 轉換為 1年1班
    }
    return '';
}

downloadExcelBtn.addEventListener('click', exportToExcel);

// 處理匯入的學生資料
function processImportedData(data) {
    if (data.length <= 1) {
        alert('匯入的檔案中沒有學生資料，請確認格式！');
        return;
    }

    students = {};
    const headerRow = data[0];
    const studentRows = data.slice(1);

    const idIndex = headerRow.indexOf('座號');
    const nameIndex = headerRow.indexOf('姓名');
    const groupIndex = headerRow.indexOf('組別');
    const seatIndex = headerRow.indexOf('座位');

    const missingHeaders = [];
    if (idIndex === -1) missingHeaders.push('座號');
    if (nameIndex === -1) missingHeaders.push('姓名');
    if (groupIndex === -1) missingHeaders.push('組別');
    if (seatIndex === -1) missingHeaders.push('座位');

    if (missingHeaders.length > 0) {
        alert(`匯入的 Excel 檔案標題不正確，缺少以下欄位：${missingHeaders.join('、')}。請確保第一列包含「座號」、「姓名」、「組別」、「座位」這四個欄位。`);
        return;
    }
    
    const studentData = studentRows.filter(row => {
        const seatValue = String(row[seatIndex]).trim();
        const isValidSeatFormat = /^\d+-\d+$/.test(seatValue);
        
        if (!row[idIndex] || !row[nameIndex] || !row[groupIndex] || !row[seatIndex]) {
            return false;
        }
        
        if (!isValidSeatFormat) {
            console.error(`座位格式不正確：${seatValue}`);
            return false;
        }

        return !isNaN(parseInt(row[groupIndex])) && !isNaN(parseInt(seatValue.split('-')[0])) && !isNaN(parseInt(seatValue.split('-')[1]));
    });
    
    if (studentData.length === 0) {
        alert('匯入的學生名單是空的，或所有資料的「座位」欄位格式不正確。請確保「座位」欄位格式為「組別-座位編號」，例如：「1-1」。');
        return;
    }

    studentData.forEach(row => {
        const student = {
            id: row[idIndex],
            name: row[nameIndex],
            group: parseInt(row[groupIndex]),
            seat: String(row[seatIndex]).trim(),
            plusScore: 0,
            projectGrade: '作品',
            projectModifier: '無',
            projectScore: 0,
            totalScore: 0
        };
        students[student.seat] = student;
    });

    currentClassName = classNameInput.value || '未命名班級';
    alert(`成功匯入 ${studentData.length} 位學生名單到 ${currentClassName}！`);
    
    createSeatingChart();
    updateSummary();
}

// 建立座位表的主要函式
function createSeatingChart() {
    seatingChartGrid.innerHTML = '';
    const groups = {};
    
    Object.values(students).forEach(student => {
        if (!groups[student.group]) {
            groups[student.group] = {};
        }
        groups[student.group][student.seat] = student;
    });

    const groupOrder = [4, 5, 6, 1, 2, 3];
    groupOrder.forEach(i => {
        const groupElement = document.createElement('div');
        groupElement.classList.add('group');
        groupElement.id = `group${i}`;
        
        const chineseNumeral = chineseNumerals[i] || i;
        groupElement.innerHTML = `<h3 class="group-title">第 ${chineseNumeral} 組</h3>`;
        
        const groupStudents = groups[i] || {};
        
        const fullSeatList = [
            `${i}-1`, `${i}-2`, `${i}-3`,
            `${i}-4`, `${i}-5`, `${i}-6`
        ];
        
        fullSeatList.forEach(seatKey => {
            const student = groupStudents[seatKey];
            const seatElement = document.createElement('div');
            seatElement.classList.add('seat');
            
            if (student) {
                const gradeOptions = ['作品', 'A', 'B', 'C', 'D', 'E', '假'].map(grade => {
                    return `<option value="${grade}" ${student.projectGrade === grade ? 'selected' : ''}>${grade}</option>`;
                }).join('');
                
                const modifierOptions = ['無', '+', '-'].map(mod => {
                    return `<option value="${mod}" ${student.projectModifier === mod ? 'selected' : ''}>${mod}</option>`;
                }).join('');

                seatElement.innerHTML = `
                    <div class="seat-row seat-row-1">
                        <span class="student-name" data-seat="${student.seat}">${student.name}</span>
                        <span class="student-id">座號 ${student.id}</span>
                    </div>
                    <div class="seat-row seat-row-2">
                        <div class="project-score-group">
                            <select id="project-grade-select-${student.seat}" onchange="updateProjectGrade('${student.seat}')">
                                ${gradeOptions}
                            </select>
                            <select id="project-modifier-select-${student.seat}" onchange="updateProjectModifier('${student.seat}')">
                                ${modifierOptions}
                            </select>
                        </div>
                    </div>
                    <div class="seat-row seat-row-3">
                        <div class="plus-score">加分總: <span id="plus-score-display-${student.seat}">${student.plusScore}</span></div>
                        <div class="score-control">
                            <button class="up-arrow" onclick="incrementPlusScore('${student.seat}')">▲</button>
                            <button class="down-arrow" onclick="decrementPlusScore('${student.seat}')">▼</button>
                        </div>
                    </div>
                `;
            } else {
                seatElement.innerHTML = `<span class="student-name">空位</span>`;
            }
            groupElement.appendChild(seatElement);
        });
        
        seatingChartGrid.appendChild(groupElement);
    });
}

// 更新作品評級
function updateProjectGrade(seatId) {
    const select = document.getElementById(`project-grade-select-${seatId}`);
    students[seatId].projectGrade = select.value;
    updateStudentTotalScore(seatId);
    updateSummary();
}

// 更新作品調整分數
function updateProjectModifier(seatId) {
    const select = document.getElementById(`project-modifier-select-${seatId}`);
    students[seatId].projectModifier = select.value;
    updateStudentTotalScore(seatId);
    updateSummary();
}

// 增加「加分總」分數
function incrementPlusScore(seatId) {
    if (!students[seatId]) {
        return;
    }
    students[seatId].plusScore += 1;
    updatePlusScoreDisplay(seatId);
    updateStudentTotalScore(seatId);
    updateSummary();
}

// 減少「加分總」分數 (已允許負分)
function decrementPlusScore(seatId) {
    if (!students[seatId]) {
        return;
    }
    students[seatId].plusScore -= 1;
    updatePlusScoreDisplay(seatId);
    updateStudentTotalScore(seatId);
    updateSummary();
}

// 更新單一學生總分
function updateStudentTotalScore(seatId) {
    const student = students[seatId];
    const gradeScore = projectGradeScores[student.projectGrade] || 0;
    const modifierScore = projectModifierScores[student.projectModifier] || 0;
    student.projectScore = gradeScore + modifierScore;
    student.totalScore = student.plusScore + student.projectScore;
}

// 更新「加分總」顯示
function updatePlusScoreDisplay(seatId) {
    const student = students[seatId];
    document.getElementById(`plus-score-display-${student.seat}`).textContent = student.plusScore;
}

// 更新成績總覽區塊
function updateSummary() {
    scoreSummary.innerHTML = '';
    const sortedStudents = Object.values(students).sort((a, b) => a.id - b.id);
    sortedStudents.forEach(student => {
        const p = document.createElement('p');
        const modifierText = student.projectModifier === '無' ? '' : student.projectModifier;
        p.textContent = `第 ${student.group} 組 - ${student.name}: 加分總 ${student.plusScore} + 作品分數 (${student.projectGrade}${modifierText}) ${student.projectScore} = 總分 ${student.totalScore}`;
        scoreSummary.appendChild(p);
    });
}

// 匯出資料到 Excel 檔案 
function exportToExcel() {
    const data = Object.values(students).map(student => {
        const modifierText = student.projectModifier === '無' ? '' : student.projectModifier;
        return {
            班級: currentClassName,
            組別: student.group,
            座號: student.id,
            姓名: student.name,
            加分總: student.plusScore,
            作品分數: `${student.projectGrade}${modifierText}`,
            總分: student.totalScore
        };
    });

    if (data.length === 0) {
        alert('沒有學生資料可以下載！');
        return;
    }
    
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, '學生總成績');
    
    // 取得當前日期並格式化
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const day = String(today.getDate()).padStart(2, '0');
    const dateString = `${year}-${month}-${day}`;
    
    // 將日期加入檔名
    const fileName = `${currentClassName}_學生總成績_${dateString}.xlsx`;
    XLSX.writeFile(workbook, fileName);
}